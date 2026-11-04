#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
给 PDF 添加页眉和“当前页 / 总页数”页码（页码默认右下角，页眉默认顶部）。

页码格式（v1.4 起与公文排版一致）：
    plain  “1”
    dash   “— 1 —”
    total  “1 / n”
对齐方式：left / center / right。left 和 right 用同一个水平边距，
center 固定在页面水平正中、忽略水平边距（与 AHK 界面的口径一致）。

页眉（v1.13 新增）：
    --header 为空串时，自动用“输入 PDF 的主文件名（不含扩展名）”；
    --header-align 三选：left / center / right；
    --header-margin-top 是页眉基线距页面上边的距离（cm），默认 1.2；
    --header-size 默认与 --font-size 相同。

两个开关（v1.15 新增，对应界面上的「页码格式：无」和「添加页眉」复选框）：
    --no-page-numbers  只画页眉、不画页码；
    --no-header        只画页码、不画页眉。
    注意「页眉文字留空」在后端是“用文件名”，不能拿来表示“不加页眉”，
    所以“不加页眉”必须用独立的 --no-header 表达，否则会意外加上文件名页眉。
    两个开关同时给时什么也不画，但仍然会清理上一次的旧叠加层（等于“撤销”）。
    页眉文字通常是中文，必须用 Windows 系统字体（msyh.ttc 等）注册进 reportlab，
    否则 reportlab 的内置 Helvetica 只有拉丁字形，中文会整段丢失。

实现原理（重要，不直观）：
    1. 用 reportlab 生成一个“页眉 + 页码”的透明单页 PDF（白纸其实没有背景）；
    2. 用 pypdf 把这一页叠到原 PDF 每一页的最上层（merge_page）。
    因为只在最上层叠加、完全不碰原内容流，所以：
      - 原 PDF 的字体、图片、分页、表格一律不变（不会像 Word 转换那样重排版面）；
      - 扫描件（整页是图片）也能加页码；
      - 原文件不动，输出到新文件。
    代价是需要 pypdf + reportlab 两个纯 Python 库。

    叠加层里还会画一个 0.1pt、不可见（渲染模式 3）的标记文本 “DocFmtPN”。
    它不影响显示，只用于重跑时定位并整块删除“上一次由本工具加的叠加层”，
    详见 _strip_our_overlay_ops()。

被 PDF 加页码工具.ahk 调用，也可以单独在命令行用：
    python pdf_page_number.py --out 输出.pdf 输入.pdf
    python pdf_page_number.py --out 输出.pdf --start 5 --font-size 12 输入.pdf
    python pdf_page_number.py --out 输出.pdf --format dash --align center 输入.pdf
    python pdf_page_number.py --out 输出.pdf --header "某某单位" --header-align center 输入.pdf
    python pdf_page_number.py --out 输出.pdf --no-header 输入.pdf
    python pdf_page_number.py --out 输出.pdf --no-page-numbers 输入.pdf

退出码：0 = 成功；2 = 失败（原因写在 --status 指定的文件里）。
"""

from __future__ import annotations

import argparse
import io
import os
import re
import shutil
import sys

# A4 纵向，单位 pt（1pt = 1/72 英寸）。仅在页面本身缺 /MediaBox 时兜底。
A4_WIDTH = 595.276
A4_HEIGHT = 841.89

CM = 72.0 / 2.54  # 1 厘米 = 28.35 pt

# 叠加层的“指纹”：一个只有平移、不画任何东西的 cm 变换。
# 为什么不用文字当标记：文字即使设成“不绘制”（渲染模式 3），仍会被文本提取、
# 搜索、复制粘贴读到，用户复制 PDF 文字时就会多出一串莫名其妙的字符。
# 纯 cm 变换既不影响显示、也不会被提取，6 个数字又足够独特（1.13/1.31 是版本号暗示）。
# 位置固定为 (1.13pt, 1.31pt)，几乎不可能与真实内容撞车；
# 重跑时按它定位整张 overlay（见 _strip_our_overlay_ops）。
OVERLAY_MARKER_MATRIX = (1.0, 0.0, 0.0, 1.0, 1.13, 1.31)

# v1.17：页眉、页码各用一个独立指纹，方便“只清页眉保留页码”或反过来。
#   · 1.13/1.31 继续代表 v1.13–v1.16 的“合并层”（页眉+页码在同一外层 q 里）；
#   · 1.13/1.32、1.13/1.33 分别是新版的页眉块、页码块。
# 三个指纹都只做平移、不画东西，对显示和文字提取都没有影响。
OVERLAY_HEADER_MARKER = (1.0, 0.0, 0.0, 1.0, 1.13, 1.32)
OVERLAY_PAGE_MARKER = (1.0, 0.0, 0.0, 1.0, 1.13, 1.33)

# 识别“文字层原生页眉/页码”的搜索带：距页面视觉上/下边各 4cm。
# 公文页脚通常 1.2–2.5cm，选 4cm 能覆盖绝大多数页边距，又不会深入正文。
NATIVE_BAND_CM = 4.0
NATIVE_BAND_PT = NATIVE_BAND_CM * CM

# 中文字体候选：先 Windows 自带的微软雅黑，再宋体、黑体、楷体。
# .ttc 是字体集合，必须给 subfontIndex（0 = 集合里的第一个字重），
# 少了这个参数 reportlab 会报 “not a TrueType font” 或直接读错偏移。
CJK_FONT_CANDIDATES = [
    (r"C:\Windows\Fonts\msyh.ttc", 0),
    (r"C:\Windows\Fonts\msyh.ttf", 0),
    (r"C:\Windows\Fonts\simsun.ttc", 0),
    (r"C:\Windows\Fonts\simhei.ttf", 0),
    (r"C:\Windows\Fonts\simkai.ttf", 0),
]
CJK_FONT_NAME = "DocFmtCJK"
_CJK_FONT_CACHE = {}   # {"font": 字体名或 None, "err": 原因}，进程内只探测一次


class PdfNumberError(Exception):
    """业务错误：把中文原因带给 AHK 显示。"""


def get_inherited(page, *keys, default=None):
    """
    读取页面属性，并沿 /Parent 向上找继承值。

    PDF 允许 /MediaBox、/CropBox、/Rotate 写在父节点（Pages）上，由所有子页面继承。
    pypdf 的 page.mediabox / page.rotation 只查页面自身，直接用它会在少数文件上
    取到默认 A4 或 0°，导致页码位置算错，所以这里自己实现继承查找。
    """
    node = page
    for _ in range(32):  # 防御异常 PDF 的 /Parent 环
        if node is None:
            break
        for key in keys:
            try:
                if key in node:
                    return node[key].get_object()
            except (TypeError, AttributeError):
                continue
        try:
            node = node.get("/Parent")
        except AttributeError:
            break
        node = node.get_object() if node is not None else None
    return default


def rect_of(value, fallback):
    """把 /MediaBox 之类的数组转成 RectangleObject，异常时用 fallback。"""
    from pypdf.generic import RectangleObject

    if value is None:
        return RectangleObject(fallback)
    try:
        rect = RectangleObject(value)
        if float(rect.width) <= 0 or float(rect.height) <= 0:
            return RectangleObject(fallback)
        return rect
    except Exception:
        return RectangleObject(fallback)


def page_geometry(page):
    """
    返回 (mediabox, cropbox, rotation)。

    显示时以 CropBox 为准，所以页码的“右边距”“下边距”都按 CropBox 计算。
    merge_page 也按 CropBox 裁剪叠加层，两边保持一致才不会错位。
    """
    mb = rect_of(get_inherited(page, "/MediaBox"), [0, 0, A4_WIDTH, A4_HEIGHT])
    cb = rect_of(get_inherited(page, "/CropBox"), list(mb))
    raw_rot = get_inherited(page, "/Rotate", default=0)
    try:
        rot = int(round(float(raw_rot) / 90.0)) * 90 % 360
    except (TypeError, ValueError):
        rot = 0
    return mb, cb, rot


def visual_to_content(xv, yv, cropbox, rotation):
    """
    把“视觉坐标”（屏幕上看到的页面左下角为原点）换算成 PDF 内容坐标。

    为什么需要这一步：/Rotate 只是让阅读器把已经画好的内容整体旋转显示，
    内容流本身的坐标系并没有变。merge_page 直接在内容坐标系里拼接，
    所以想让页码出现在“视觉上的右下角”，必须先做逆旋转。

    设 CropBox 宽 W、高 H（内容坐标系），视觉坐标 (xv, yv)：
        rot = 0   : xr = xv,        yr = yv
        rot = 90  : xr = W - yv,    yr = xv        （内容顺时针转 90° 显示）
        rot = 180 : xr = W - xv,    yr = H - yv
        rot = 270 : xr = yv,        yr = H - xv
    最后加上 CropBox 左下角的偏移，得到绝对内容坐标。
    """
    w, h = float(cropbox.width), float(cropbox.height)
    if rotation == 90:
        return float(cropbox.left) + (w - yv), float(cropbox.bottom) + xv
    if rotation == 180:
        return float(cropbox.left) + (w - xv), float(cropbox.bottom) + (h - yv)
    if rotation == 270:
        return float(cropbox.left) + yv, float(cropbox.bottom) + (h - xv)
    return float(cropbox.left) + xv, float(cropbox.bottom) + yv


def register_cjk_font():
    """
    把系统里的中文字体注册进 reportlab，返回 (字体名 或 None, 失败原因)。

    为什么必须注册：reportlab 内置的 Helvetica 只有拉丁字形，drawString 画中文
    不会报错，但输出里这些字根本没有字形（Glyph 0），看起来就是“字丢了”。
    TTFont 会把字体子集嵌进生成的叠加层；最后 merge 进原 PDF，所以只需嵌
    用到的十几个字，体积通常只有几十 KB。

    .ttc 是“字体集合”，必须显式给 subfontIndex：
        TTFont(name, path, subfontIndex=0)
    少了它 reportlab 会把它当成普通 .ttf 去读，偏移全错。

    任何候选都失败时返回 (None, 原因)，调用方负责决定是继续（中文会缺字）
    还是把原因写进 --status 让用户看到；这里不抛异常，避免因为字体问题
    连页码都加不上。
    """
    if "font" in _CJK_FONT_CACHE:
        return _CJK_FONT_CACHE["font"], _CJK_FONT_CACHE["err"]

    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    tried = []
    for path, sub_idx in CJK_FONT_CANDIDATES:
        if not os.path.isfile(path):
            tried.append(os.path.basename(path) + "（不存在）")
            continue
        try:
            pdfmetrics.registerFont(TTFont(CJK_FONT_NAME, path, subfontIndex=sub_idx))
            _CJK_FONT_CACHE["font"] = CJK_FONT_NAME
            _CJK_FONT_CACHE["err"] = ""
            return CJK_FONT_NAME, ""
        except Exception as exc:  # 字体损坏 / 被占用 / 格式不兼容
            tried.append("%s（%s）" % (os.path.basename(path), exc))

    _CJK_FONT_CACHE["font"] = None
    _CJK_FONT_CACHE["err"] = "未找到可用的中文字体：" + "；".join(tried)
    return None, _CJK_FONT_CACHE["err"]


def _draw_text_line(c, text, font_name, font_size, x, y, rotation, align,
                    marker=None):
    """
    在 overlay 上画一行文字。

    x / y 是绝对内容坐标（定位点基线）；rotation 是“让文字在旋转页面上看起来
    仍然水平”所需的文字旋转角度（等于页面 /Rotate，推导见 visual_to_content）。

    align 决定定位点含义：left = 文字左端、center = 文字中心、right = 文字右端。
    必须在 translate 之后的局部坐标系里调用对应的 drawXxxString(0, 0)，不能
    先算文字宽度再自己减——reportlab 的 *String 系列会按当前字体/字号算宽度，
    中文字符和旋转后的坐标系都能正确处理。

    每行自己 saveState/restoreState：translate 是累积的，
    否则第二行的坐标会叠加到第一行上。

    v1.17：marker 是本行的“指纹 cm”（见 OVERLAY_*_MARKER），必须写在
    restoreState() 之前，这样它才在这一行自己的 q/Q 块里面；清理时按指纹
    找到这个 q 块，就能只删页眉或只删页码。
    """
    c.saveState()
    c.translate(x, y)
    if rotation:
        c.rotate(rotation)  # reportlab 的 rotate 是逆时针
    c.setFillColorRGB(0, 0, 0)
    c.setFont(font_name, font_size)
    if align == "left":
        c.drawString(0, 0, text)
    elif align == "center":
        c.drawCentredString(0, 0, text)
    else:
        # 右对齐：文字向旋转后坐标系的 -x 方向延伸，视觉上就是向左，右端对齐到 margin
        c.drawRightString(0, 0, text)
    if marker is not None:
        # 纯平移，不画任何东西；放在文字之后、restoreState 之前。
        c.transform(*marker)
    c.restoreState()


def make_overlay(width, height, text, font_name, font_size, x, y, rotation, align="right",
                 header_text="", header_font_name="", header_size=0.0,
                 header_x=0.0, header_y=0.0, header_align="left"):
    """
    生成只含“页眉 + 页码”的透明单页 PDF，并返回对应的 PageObject。

    v1.13 起页眉与页码画在同一张 overlay 上、只 merge 一次；v1.17 改成
    “每行一个独立 q/Q 块 + 各自指纹”，不再套最外层 q/Q。这样：
      · 只清页眉时，按 OVERLAY_HEADER_MARKER 删掉页眉块，页码块原样保留；
      · 只清页码时同理；
      · v1.13–v1.16 的旧合并层由 _strip_legacy_combined_ops() 按位置拆分。
    """
    from pypdf import PdfReader
    from reportlab.pdfgen import canvas

    if not header_size:
        header_size = font_size
    if not header_font_name:
        header_font_name = font_name

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height), pageCompression=1)
    drew = False
    if header_text:
        _draw_text_line(c, header_text, header_font_name, header_size,
                        header_x, header_y, rotation, header_align,
                        marker=OVERLAY_HEADER_MARKER)
        drew = True
    if text:
        _draw_text_line(c, text, font_name, font_size, x, y, rotation, align,
                        marker=OVERLAY_PAGE_MARKER)
        drew = True
    if not drew:
        # 两样都不画时连空 overlay 都不生成，避免留一层只有指纹的空内容流。
        return None
    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def format_page_text(number, total, style):
    """
    按格式拼页码文字。与公文排版的三种格式一一对应：
        plain  “1”
        dash   “— 1 —”
        total  “1 / n”
    这里用半角空格，和 Word 页脚里的“— 1 —”观感一致。
    """
    if style == "plain":
        return f"{number}"
    if style == "dash":
        return f"— {number} —"
    if style == "total":
        return f"{number} / {total}"
    raise PdfNumberError(f"未知的页码格式：{style}")


def _is_marker_cm(operands, marker=OVERLAY_MARKER_MATRIX):
    """判断一个 cm 操作是不是指定指纹（默认是 v1.13–v1.16 的合并层指纹）。"""
    if len(operands) != 6:
        return False
    try:
        values = [float(v) for v in operands]
    except (TypeError, ValueError):
        return False
    return all(abs(v - want) < 1e-3 for v, want in zip(values, marker))


def _find_marker_index(ops, marker=OVERLAY_MARKER_MATRIX):
    """从内容流末尾往前找指定指纹 cm 变换，返回下标；没有则 None。"""
    for i in range(len(ops) - 1, -1, -1):
        operands, operator = ops[i]
        if operator == b"cm" and _is_marker_cm(operands, marker):
            return i
    return None


def _has_any_overlay_marker(ops):
    """内容流里是否还留着本工具任意版本的指纹。"""
    for operands, operator in ops:
        if operator == b"cm" and (
            _is_marker_cm(operands, OVERLAY_MARKER_MATRIX)
            or _is_marker_cm(operands, OVERLAY_HEADER_MARKER)
            or _is_marker_cm(operands, OVERLAY_PAGE_MARKER)
        ):
            return True
    return False


def _enclosing_q_block(ops, index):
    """
    返回包含 index 的“最内层” q ... Q 块的 (start, end) 下标；找不到返回 None。

    做法：先扫一遍所有 q/Q，用栈配对成 (start, end)，再按 start 从大到小
    （越靠后越内层）挑第一个真正把 index 包在中间的块。
    v1.17 的页眉块/页码块各自一对 q/Q，指纹写在自己块的末尾，所以这里拿到的
    就是那一行文字的边界，不会把另一行包进去。
    """
    stack = []
    pairs = []
    for i, (_, operator) in enumerate(ops):
        if operator == b"q":
            stack.append(i)
        elif operator == b"Q" and stack:
            pairs.append((stack.pop(), i))
    pairs.sort(key=lambda p: p[0], reverse=True)
    for start, end in pairs:
        if start < index < end:
            return start, end
    return None


def _strip_marked_ops(ops, marker, max_layers=4):
    """
    删除内容流末尾由本工具某个指纹标记的 q/Q 块。

    max_layers 限制最多删几层，防止内容流被别的工具改坏后死循环。
    v1.17 的页眉块和页码块是分开的，分别用 OVERLAY_HEADER_MARKER /
    OVERLAY_PAGE_MARKER 调用本函数即可只删其中一类。
    """
    ops = list(ops)
    removed = 0
    while removed < max_layers:
        marker_idx = _find_marker_index(ops, marker)
        if marker_idx is None:
            break
        block = _enclosing_q_block(ops, marker_idx)
        if block is None:
            # 指纹在，但找不到所属的 q 块（内容流不完整）：放弃清理，保持原样
            break
        start, end = block
        del ops[start:end + 1]
        removed += 1
    return ops


def _q_pairs(ops):
    """把内容流里的 q/Q 配对成 (start, end) 列表；不完整的 q 忽略。"""
    stack = []
    pairs = []
    for i, (_, operator) in enumerate(ops):
        if operator == b"q":
            stack.append(i)
        elif operator == b"Q" and stack:
            pairs.append((stack.pop(), i))
    return pairs


def _inner_text_blocks(ops, outer_start, outer_end):
    """
    找出 outer q 块内部“真正画了字”的最内层 q 块。

    返回 [{start, end, visual_y}]，按 start 升序。
    v1.13–v1.16 的合并 overlay 结构固定为：
        q                       ← 外层
          q cm … BT … Tj … ET Q  ← 页眉
          q cm … BT … Tj … ET Q  ← 页码
          cm(指纹)
        Q
    这里只取“叶子”q 块（内部不再包含别的 q 块），避免把外层重复算进去。
    """
    pairs = _q_pairs(ops)
    candidates = []
    for start, end in pairs:
        if not (outer_start < start and end < outer_end):
            continue
        if not any(op in (b"Tj", b"TJ") for _, op in ops[start:end + 1]):
            continue
        candidates.append((start, end))

    # 去掉包含其他候选的块，只留最内层
    leaves = []
    for start, end in candidates:
        if any(s != start and start < s and e < end for s, e in candidates):
            continue
        leaves.append((start, end))

    result = []
    for start, end in sorted(leaves):
        position = _block_translation(ops, start, end)
        result.append({"start": start, "end": end, "position": position})
    return result


def _block_translation(ops, start, end):
    """取 q 块内第一个 cm 的 (e, f) 平移；没有则 None。"""
    for operands, operator in ops[start:end + 1]:
        if operator == b"cm" and len(operands) == 6:
            try:
                return float(operands[4]), float(operands[5])
            except (TypeError, ValueError):
                return None
    return None


def _visual_y_from_content(x, y, cropbox, rotation):
    """
    visual_to_content() 的逆运算：把内容坐标换算成“视觉上距页面下边的高度”。

    只关心 y（用来判断页眉带/页脚带）。公式与 visual_to_content 一一对应：
        rot = 0   : vy = y  - bottom
        rot = 90  : vy = W  - (x - left)
        rot = 180 : vy = H  - (y - bottom)
        rot = 270 : vy = x  - left
    """
    w, h = float(cropbox.width), float(cropbox.height)
    left, bottom = float(cropbox.left), float(cropbox.bottom)
    if rotation == 90:
        return w - (x - left)
    if rotation == 180:
        return h - (y - bottom)
    if rotation == 270:
        return x - left
    return y - bottom


def _strip_legacy_combined_ops(ops, keep_header, keep_page, cropbox, rotation,
                               page_height):
    """
    处理 v1.13–v1.16 的“合并层”（一个外层 q 里同时画页眉和页码）。

    v1.17 只清其中一项时不能整块删除，否则会把要保留的那一项也删掉。
    这里按每个文字块的视觉 y 把页眉/页码分开，只删用户要求清掉的部分：
      · 两项都保留：原样不动；
      · 两项都清空：删除整个外层 q 块；
      · 只清页码：删除下面的文字块，保留上面的页眉；
      · 只清页眉：删除上面的文字块，保留下面的页码。
    删完若块内已经没有文字（只剩指纹），连外层 q 块一起删掉。
    """
    ops = list(ops)
    if keep_header and keep_page:
        return ops

    # 先收集所有合并层的外层 q 块。按 start 从大到小处理，
    # 这样前面块的索引不会被后面块的增删影响。
    blocks = []
    seen = set()
    for i, (operands, operator) in enumerate(ops):
        if operator == b"cm" and _is_marker_cm(operands, OVERLAY_MARKER_MATRIX):
            outer = _enclosing_q_block(ops, i)
            if outer and outer not in seen:
                seen.add(outer)
                blocks.append(outer)
    for start, end in sorted(blocks, key=lambda p: p[0], reverse=True):
        block = ops[start:end + 1]
        inner = _inner_text_blocks(block, 0, len(block) - 1)
        for item in inner:
            if item["position"] is None:
                item["visual_y"] = -1.0
            else:
                x, y = item["position"]
                item["visual_y"] = _visual_y_from_content(x, y, cropbox, rotation)

        if not inner:
            # 只剩指纹的空壳，直接删掉
            ops[start:end + 1] = []
            continue

        if not keep_header and not keep_page:
            ops[start:end + 1] = []
            continue

        # 分类：两块及以上时，y 最大=页眉、y 最小=页码；只有一块时按页面上/下半部判断。
        if len(inner) >= 2:
            ordered = sorted(inner, key=lambda b: b["visual_y"])
            page_blocks = [ordered[0]]
            header_blocks = [ordered[-1]]
        else:
            only = inner[0]
            if only["visual_y"] >= page_height / 2.0:
                header_blocks = [only]
                page_blocks = []
            else:
                page_blocks = [only]
                header_blocks = []

        delete_blocks = []
        if not keep_header:
            delete_blocks.extend(header_blocks)
        if not keep_page:
            delete_blocks.extend(page_blocks)
        if not delete_blocks:
            continue

        # 从后往前删，索引不会串位
        for item in sorted(delete_blocks, key=lambda b: b["start"], reverse=True):
            del block[item["start"]:item["end"] + 1]

        # 删完若已经没有文字，连外层 q 块和指纹一起删
        if not any(op in (b"Tj", b"TJ") for _, op in block):
            ops[start:end + 1] = []
        else:
            ops[start:end + 1] = block

    return ops


def _strip_our_overlay_ops(ops, max_layers=4):
    """兼容旧调用：删除 v1.13–v1.16 的整个合并层（两项都清空时使用）。"""
    return _strip_marked_ops(ops, OVERLAY_MARKER_MATRIX, max_layers=max_layers)


# 旧版（v1.10–v1.12）页码文字里可能出现各种破折号：
#   · PDF 内容流里是 WinAnsi 字节 0x97，pypdf 可能把它解码成 U+2014、U+0160 等；
#   · 用户也可能用普通的 “-” 或全角 “－”。
# 所以这里不直接比字符串，而是先归一化：只保留数字和斜杠，破折号/空格全部丢掉，
# 出现任何其他字符（字母、汉字）就判定“不是页码”，宁可少删。
_LEGACY_DASH_CHARS = set(" \t\u00a0-\u2010\u2011\u2012\u2013\u2014\u2015\u2212\u30fc\u0160\uff0d\uff0f")


def _normalize_legacy_text(value):
    """
    把 Tj 操作数归一化成“纯数字 / 数字+斜杠”，不是页码形状时返回 None。

    pypdf 解析出的操作数可能是 str（已按字体编码解码）或 bytes（原样字节），
    两种都要处理；解不出来就当不是页码。
    """
    if isinstance(value, bytes):
        text = value.decode("cp1252", "ignore")
    elif isinstance(value, str):
        text = value
    else:
        return None
    kept = []
    for ch in text:
        if ch.isdigit() or ch == "/":
            kept.append(ch)
        elif ch in _LEGACY_DASH_CHARS:
            continue
        else:
            return None
    return "".join(kept)


def _legacy_text_shape_ok(text):
    """归一化后的文本必须是 1、12、1/3 这类页码形状。"""
    if not text:
        return False
    if "/" in text:
        parts = text.split("/")
        return len(parts) == 2 and all(p.isdigit() and p for p in parts)
    return text.isdigit()


def _text_blocks(ops):
    """把内容流里的 BT ... ET 文本对象配对成 (start, end) 列表。"""
    blocks = []
    start = None
    for i, (_, operator) in enumerate(ops):
        if operator == b"BT":
            start = i
        elif operator == b"ET" and start is not None:
            blocks.append((start, i))
            start = None
    return blocks


def _legacy_block_position(ops, start, lookback=12):
    """
    找文本块前面最近的一个 cm 平移，返回 (e, f)；找不到返回 None。

    旧版 overlay 的画法固定是：q → cm(平移到页码位置) → rg → BT/Tf … ET → Q，
    所以“紧挨在文本块前面”的 cm 就是页码的落点。正文一般用 Tm 定位而不是 cm，
    这一步等于要求“这个文本块确实是被平移到页边的那一行”，可以再挡掉一批误判。
    """
    for i in range(start - 1, max(-1, start - 1 - lookback), -1):
        operands, operator = ops[i]
        if operator == b"cm" and len(operands) == 6:
            try:
                return float(operands[4]), float(operands[5])
            except (TypeError, ValueError):
                return None
    return None


def _nearest_tf_size(ops, index, lookback=20):
    """
    从 index 往前找最近的 Tf，返回字号；找不到返回 None。

    为什么需要：reportlab 画一行字的操作序列是
        BT /F1 10.5 Tf 12.6 TL ET      ← 只设字体，不画字
        BT 1 0 0 1 0 0 Tm (文字) Tj T* ET   ← 真正画字的那块，里面没有 Tf
    所以“画字块”自己往往查不到字号，必须往前找同一个文本对象之前的 Tf。
    """
    for i in range(index, max(-1, index - lookback), -1):
        operands, operator = ops[i]
        if operator == b"Tf" and len(operands) >= 2:
            try:
                return float(operands[1])
            except (TypeError, ValueError):
                return None
    return None


def _looks_like_legacy_page_number(seg, font_sizes, position, cropbox_bottom, page_height,
                                   fallback_size=None):
    """
    判断一个 BT..ET 文本块是不是旧版（无指纹）页码叠加层。

    条件（全部满足才算）：
      1. 只画了一次文字（Tj），且不是 TJ 数组——reportlab 画一行字就是这个形状；
      2. 块内所有 Tf 字号都等于给定的字号之一（页码字号 / 页眉字号）；
      3. 文字归一化后严格是页码形状（1、12、1/3）；
      4. 落点在页脚或页眉带内（距 CropBox 下边 / 上边 100pt 以内）。
    正文里偶尔有一行只有数字也会命中前三条，但第 4 条要求它被平移到页面
    最上/最下 3.5cm 以内，正常正文不会出现在那里。
    """
    texts = []
    sizes = []
    for operands, operator in seg:
        if operator == b"Tf" and len(operands) >= 2:
            try:
                sizes.append(float(operands[1]))
            except (TypeError, ValueError):
                return False
        elif operator == b"Tj" and operands:
            texts.append(operands[0])
        elif operator == b"TJ":
            return False       # TJ 是字距数组，旧版页码不会这么画，直接判不是
    if len(texts) != 1:
        return False
    if not sizes:
        # 画字块里没有 Tf：用前面最近的 Tf 字号（见 _nearest_tf_size 的说明）
        if fallback_size is None:
            return False
        sizes = [fallback_size]
    if not all(any(abs(s - want) < 0.01 for want in font_sizes) for s in sizes):
        return False
    if not _legacy_text_shape_ok(_normalize_legacy_text(texts[0])):
        return False
    if position is None:
        return False
    _, y = position
    return (y <= cropbox_bottom + 100) or (y >= cropbox_bottom + page_height - 100)


def _strip_legacy_overlay_ops(ops, font_sizes, cropbox_bottom=0.0, page_height=0.0,
                              max_blocks=3):
    """
    兜底清理 v1.10–v1.12 留下的旧页码层（那些版本没有写指纹）。

    只从内容流末尾往前处理：跳过没有画字的“字体设置脚手架块”，
    遇到第一块不像旧页码的就立刻停止，最多删 max_blocks 块。
    这样既能把之前几次叠上去的旧页码清掉，又不会像 v1.10 那样
    按字号扫全文、误删 10.5pt（五号）的正文。
    """
    ops = list(ops)
    removed = 0
    while removed < max_blocks:
        blocks = _text_blocks(ops)
        if not blocks:
            break
        deleted = False
        for start, end in reversed(blocks):
            seg = ops[start:end + 1]
            if not any(op in (b"Tj", b"TJ") for _, op in seg):
                continue       # 只有 Tf/TL 的脚手架块，继续往前找真正画字的那块
            position = _legacy_block_position(ops, start)
            if not _looks_like_legacy_page_number(
                    seg, font_sizes, position, cropbox_bottom, page_height,
                    fallback_size=_nearest_tf_size(ops, start)):
                break          # 碰到正文（或其它内容）就停，绝不深入
            del ops[start:end + 1]
            removed += 1
            deleted = True
            break
        if not deleted:
            break
    return ops


def _load_page_fonts(page):
    """
    读取页面 /Resources /Font，并转换成 pypdf 的 Font 对象。

    用于解析“文字层原生页眉/页码”的编码。拿不到资源时返回空字典，
    调用方会把这些文字当成不可解析、直接跳过（宁可漏删，不误删）。
    """
    from pypdf._page import _get_font_resources, _get_page_resources
    from pypdf._text_extraction._text_extractor import Font

    fonts = {}
    try:
        resources = _get_page_resources(page)
        font_resources = _get_font_resources(resources)
        for name, ref in font_resources.items():
            try:
                fonts[name] = Font.from_font_resource(ref.get_object())
            except Exception:
                continue
    except Exception:
        return {}
    return fonts


def _decode_text_operands(operands, extractor):
    """用当前字体把 Tj/TJ 里的操作数解码成可读文字（含 ToUnicode 映射）。"""
    from pypdf._text_extraction._text_extractor import get_text_operands

    try:
        text, _, _ = get_text_operands(
            operands, extractor.cm_matrix, extractor.tm_matrix,
            extractor.font, (0, 90, 180, 270),
        )
    except Exception:
        return ""
    # get_text_operands() 返回的是字体编码里的原始字符；pypdf 正常提取时
    # 还会过一次 character_map（ToUnicode）。这里补上同一步，否则中文会变成控制字符。
    return "".join(extractor.font.character_map.get(ch, ch) for ch in text)


def _native_text_items(ops, fonts, cropbox, rotation):
    """
    找出内容流里每个 Tj / TJ 画出的文字、视觉 y 和字号。

    只处理 Tj 和 TJ；单引号/双引号操作符（'、"）在正文里很少见，而且删除它们
    会连带丢掉换行推进，可能让后面的文字串位，所以保守跳过。

    返回 [{index, text, visual_y, font_size}]，index 是 operations 里的下标，
    调用方按这个下标删除对应的文字操作。
    """
    from pypdf._text_extraction._text_extractor import TextExtraction

    extractor = TextExtraction()
    extractor.initialize_extraction((0, 90, 180, 270), None, fonts, fonts)

    def position():
        m = _mult_matrix(extractor.tm_matrix, extractor.cm_matrix)
        return m[4], m[5]

    items = []
    for index, (operands, operator) in enumerate(ops):
        if operator == b"Tj":
            if operands:
                x, y = position()
                text = _decode_text_operands(operands[:1], extractor)
                if text:
                    items.append({
                        "index": index,
                        "text": text,
                        "visual_y": _visual_y_from_content(x, y, cropbox, rotation),
                        "font_size": extractor.font_size,
                    })
        elif operator == b"TJ":
            if operands and isinstance(operands[0], (list, tuple)):
                x, y = position()
                pieces = []
                for piece in operands[0]:
                    if isinstance(piece, (str, bytes)):
                        pieces.append(_decode_text_operands([piece], extractor))
                    elif isinstance(piece, (int, float)):
                        # 与 pypdf 保持一致：大的字距调整视为一个空格
                        if (abs(float(piece)) >= extractor._space_width * 0.95
                                and pieces and not pieces[-1].endswith(" ")):
                            pieces.append(" ")
                text = "".join(pieces)
                if text:
                    items.append({
                        "index": index,
                        "text": text,
                        "visual_y": _visual_y_from_content(x, y, cropbox, rotation),
                        "font_size": extractor.font_size,
                    })

        # 推进提取器状态（矩阵、字体、字号、当前文字缓冲）。
        if operator == b"'":
            extractor.process_operation(b"T*", [])
            extractor.process_operation(b"Tj", operands)
        elif operator == b'"' and len(operands) >= 3:
            extractor.process_operation(b"Tw", [operands[0]])
            extractor.process_operation(b"Tc", [operands[1]])
            extractor.process_operation(b"T*", [])
            extractor.process_operation(b"Tj", operands[2:])
        elif operator == b"TJ" and operands and isinstance(operands[0], (list, tuple)):
            for piece in operands[0]:
                if isinstance(piece, (str, bytes)):
                    extractor.process_operation(b"Tj", [piece])
                elif isinstance(piece, (int, float)):
                    if (abs(float(piece)) >= extractor._space_width * 0.95
                            and extractor.text and extractor.text[-1] != " "):
                        extractor.process_operation(b"Tj", [" "])
        else:
            extractor.process_operation(operator, operands)
    return items


def _mult_matrix(m, n):
    """PDF 矩阵乘法，和 pypdf 的 mult() 同义（只用于算文字落点）。"""
    a, b, c, d, e, f = m
    A, B, C, D, E, F = n
    return (
        a * A + b * C,
        a * B + b * D,
        c * A + d * C,
        c * B + d * D,
        e * A + f * C + E,
        e * B + f * D + F,
    )


_FULLWIDTH_TRANS = str.maketrans(
    "０１２３４５６７８９／－—–　，,",
    "0123456789/-—– ,,",
)


def _normalize_page_number_text(text):
    """去掉空白、统一全角字符，便于识别“1”“第 1 页”“1 / 12”等页码形状。"""
    value = str(text).translate(_FULLWIDTH_TRANS)
    return re.sub(r"\s+", "", value)


def _page_number_parts(text):
    """
    判断整段文字是不是“页码形状”，是则返回 (当前页, 总页数或 None)。

    只认整段都是页码的情况；正文里“第 3 页的正文……”这种不会命中。
    """
    value = _normalize_page_number_text(text)
    if not value:
        return None

    patterns = [
        r"^[-—–]?(\d+)[-—–]?$",
        r"^(\d+)/(\d+)$",
        r"^第(\d+)页(?:共(\d+)页)?$",
        r"^(\d+)页$",
        r"^page(\d+)(?:of(\d+))?$",
        r"^(\d+)of(\d+)$",
    ]
    for pattern in patterns:
        match = re.match(pattern, value, re.IGNORECASE)
        if match:
            first = int(match.group(1))
            second = None
            if match.lastindex and match.lastindex >= 2 and match.group(2):
                second = int(match.group(2))
            return first, second
    return None


def _text_looks_like_page_number(text):
    """整段文字是否是页码形状（不校验具体数字）。"""
    return _page_number_parts(text) is not None


def _match_native_page_number(text, expected_numbers, total):
    """
    判断一段页脚文字是不是“当前物理页”的原生页码。

    expected_numbers 是该页可接受的数字集合（物理页号 + PDF 页标签里的数字），
    这样封面不编号、正文从 1 开始的文档也能正确处理。
    形如“1 / 12”的还会校验总页数，避免把页脚的“1 / 2”之类分数误删。
    """
    parts = _page_number_parts(text)
    if parts is None:
        return False
    first, second = parts
    if first not in expected_numbers:
        return False
    if second is not None and total and second != total:
        return False
    return True


def _normalize_header_text(text):
    """页眉重复性比较用的归一化：去首尾空白、把连续空白压成一个空格。"""
    return re.sub(r"\s+", " ", str(text).strip())


def _collect_native_header_texts(pages, total):
    """
    找出“文字层里跨页重复的页眉”。

    做法：只统计视觉上距页顶 4cm 以内的文字，归一化后在同一页去重；
    在至少 60%（且不少于 2 页）的页面上出现的文字，才认为是页眉。
    这样可以避免把只在某一页出现的标题、章节名误删。

    返回候选文字集合；调用方再逐页删除命中的 Tj/TJ。
    """
    counts = {}
    for page in pages:
        try:
            _, cb, rot = page_geometry(page)
            if rot in (90, 270):
                vh = float(cb.width)
            else:
                vh = float(cb.height)
            ops = list(page.get_contents().operations)
            fonts = _load_page_fonts(page)
            items = _native_text_items(ops, fonts, cb, rot)
        except Exception:
            continue

        seen = set()
        for item in items:
            if item["visual_y"] < vh - NATIVE_BAND_PT:
                continue
            text = _normalize_header_text(item["text"])
            if not text or len(text) > 200:
                continue
            # 页顶也可能放页码（少数模板），这种不算页眉
            if _text_looks_like_page_number(text):
                continue
            seen.add(text)
        for text in seen:
            counts[text] = counts.get(text, 0) + 1

    threshold = max(2, (total * 3 + 4) // 5)  # ceil(0.6 * total)
    return {text for text, count in counts.items() if count >= threshold}


def _remove_text_items(ops, items, predicate):
    """
    删除 items 里满足 predicate 的 Tj/TJ 操作。

    只删“画字”那一个操作，保留 BT/ET、Tf、Tm 等状态操作：
    这样后面的文字位置不会因为少了一个 q/Q 或换行推进而整体串位。
    返回 (新操作列表, 删除处数)。
    """
    drop = set()
    for item in items:
        try:
            if predicate(item):
                drop.add(item["index"])
        except Exception:
            continue
    if not drop:
        return list(ops), 0
    return [op for i, op in enumerate(ops) if i not in drop], len(drop)


def _page_label_number(reader, index):
    """
    取 PDF 页标签里的数字（第 index 页，1 开始）。

    有些文档封面不编号、正文从 1 开始，物理第 2 页的页脚写的是“1”。
    这时用页标签里的数字作为第二个可接受值，就能正确识别原生页码。
    页标签不存在或是 “i / ii / iii” 这种罗马数字时返回 None。
    """
    try:
        labels = reader.page_labels
        if 0 <= index - 1 < len(labels):
            label = str(labels[index - 1])
            parts = _page_number_parts(label)
            if parts is not None:
                return parts[0]
            match = re.search(r"\d+", label)
            if match:
                return int(match.group(0))
    except Exception:
        pass
    return None


def _remove_native_text_items(ops, items, drop_header, drop_page):
    """
    按条件删除文字层里的原生页眉/页码，分别统计删除处数。

    两个条件都在同一份 items 上判断，最后一次性删除，避免删完页眉后
    页码的下标串位。返回 (新操作列表, 页眉处数, 页码处数)。
    """
    drop = set()
    header_count = 0
    page_count = 0
    for item in items:
        try:
            is_header = drop_header is not None and drop_header(item)
            is_page = drop_page is not None and drop_page(item)
        except Exception:
            continue
        if is_header:
            header_count += 1
            drop.add(item["index"])
        if is_page:
            page_count += 1
            drop.add(item["index"])
    if not drop:
        return list(ops), 0, 0
    return [op for i, op in enumerate(ops) if i not in drop], header_count, page_count


def add_page_numbers(
    src,
    dst,
    start=1,
    font_size=10.5,
    margin_right_cm=2.0,
    margin_bottom_cm=1.2,
    margin_top_cm=1.2,
    font_name="Helvetica",
    style="total",
    align="right",
    header="",
    header_align="center",
    header_size=0.0,
    page_mode="add",
    header_mode="add",
):
    """
    给 PDF 叠加页码/页眉，或按模式清空/保留。

    page_mode / header_mode 取值（v1.17）：
        add   删除本工具旧层（不动文字层原生内容），再按当前设置画新的；
        clear 删除本工具旧层，并删除文字层里能识别的原生页码/页眉，不再画；
        keep  原样保留，不删也不画。

    「清空」只处理能可靠识别的文字层内容：页脚 4cm 内的页码形状文字、
    页顶 4cm 内跨页重复的页眉文字。扫描图片里印着的页眉/页码无法删除，
    正文内容也不会因为字号相同就被误删。
    """
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import ArrayObject, ContentStream

    if page_mode not in ("add", "keep", "clear"):
        raise PdfNumberError(f"未知的页码模式：{page_mode}")
    if header_mode not in ("add", "keep", "clear"):
        raise PdfNumberError(f"未知的页眉模式：{header_mode}")
    if page_mode == "add" and style not in ("plain", "dash", "total"):
        raise PdfNumberError(f"未知的页码格式：{style}")
    if page_mode == "add" and align not in ("left", "center", "right"):
        raise PdfNumberError(f"未知的对齐方式：{align}")
    if header_mode == "add" and header_align not in ("left", "center", "right"):
        raise PdfNumberError(f"未知的页眉对齐方式：{header_align}")
    if not header_size:
        header_size = font_size

    if not os.path.isfile(src):
        raise PdfNumberError("找不到输入文件")

    try:
        reader = PdfReader(src)
    except Exception as exc:  # 损坏、不是 PDF 等
        raise PdfNumberError(f"无法读取 PDF：{exc}") from exc

    if reader.is_encrypted:
        # 只有“空密码”能解开的情况才继续；真正加密的文件不猜密码
        try:
            if not reader.decrypt(""):
                raise PdfNumberError("PDF 已加密，需要密码，无法自动加页码")
        except PdfNumberError:
            raise
        except Exception as exc:
            raise PdfNumberError(f"PDF 已加密，无法读取：{exc}") from exc

    try:
        total = len(reader.pages)
    except Exception as exc:
        raise PdfNumberError(f"无法读取页数：{exc}") from exc
    if total == 0:
        raise PdfNumberError("PDF 没有任何页面")

    margin_right = margin_right_cm * CM
    margin_bottom = margin_bottom_cm * CM
    margin_top = margin_top_cm * CM

    # 页眉字体：有中文就一定要用系统字体，否则中文会整段丢字（见 register_cjk_font）。
    # 纯英文页眉不必嵌入字体，沿用页码字体即可，输出体积更小。
    header_font_name = font_name
    font_warning = ""
    if header_mode == "add" and header and any(ord(ch) > 127 for ch in header):
        cjk_font, cjk_err = register_cjk_font()
        if cjk_font:
            header_font_name = cjk_font
            # 页码（数字）也用同一字体，避免一页里出现两种字形
            font_name = cjk_font
        else:
            font_warning = "；页眉中文无法显示（" + cjk_err + "）"

    # 清空页眉前，先跨页统计“重复出现在页顶”的文字。必须在写回之前扫原页面，
    # 否则第一页删掉的文字会影响后面页的重复性判断。
    native_header_texts = set()
    if header_mode == "clear" and total > 1:
        try:
            native_header_texts = _collect_native_header_texts(reader.pages, total)
        except Exception:
            native_header_texts = set()

    try:
        # clone_from 保留书签、表单、元数据、压缩方式，比逐页 add_page 更保真
        writer = PdfWriter(clone_from=reader)
    except Exception as exc:
        raise PdfNumberError(f"初始化输出失败：{exc}") from exc

    native_header_count = 0
    native_page_count = 0

    for index, page in enumerate(writer.pages, start=1):
        try:
            mb, cb, rot = page_geometry(page)
            # 视觉页面的宽高：90°/270° 时内容坐标系的宽高互换
            if rot in (90, 270):
                vw, vh = float(cb.height), float(cb.width)
            else:
                vw, vh = float(cb.width), float(cb.height)

            xv = vw - margin_right          # 右端距视觉页右边 margin_right
            yv = margin_bottom              # 基线距视觉页下边 margin_bottom
            # 居左/居右共用同一个水平边距；居中固定在页心，忽略水平边距。
            # 注意这里算的是“视觉坐标”，旋转页由 visual_to_content() 统一换算。
            if align == "left":
                xv = margin_right
            elif align == "center":
                xv = vw / 2.0
            x, y = visual_to_content(xv, yv, cb, rot)

            # 页眉位置：基线距“视觉页面”上边 margin_top；
            # 水平方向与页码共用同一套边距规则（居左=距左、居右=距右、居中=页心）。
            hxv = vw - margin_right
            if header_align == "left":
                hxv = margin_right
            elif header_align == "center":
                hxv = vw / 2.0
            hyv = vh - margin_top
            hx, hy = visual_to_content(hxv, hyv, cb, rot)

            contents = page.get_contents()
            if contents is not None:
                ops = list(contents.operations)
                new_ops = ops

                # ① 新版分开的指纹层：只删用户要求清掉的那一类。
                if header_mode != "keep":
                    new_ops = _strip_marked_ops(new_ops, OVERLAY_HEADER_MARKER)
                if page_mode != "keep":
                    new_ops = _strip_marked_ops(new_ops, OVERLAY_PAGE_MARKER)

                # ② v1.13–v1.16 的合并层：按文字块的上下位置拆分。
                if header_mode != "keep" or page_mode != "keep":
                    new_ops = _strip_legacy_combined_ops(
                        new_ops,
                        keep_header=(header_mode == "keep"),
                        keep_page=(page_mode == "keep"),
                        cropbox=cb,
                        rotation=rot,
                        page_height=vh,
                    )

                # ③ v1.10–v1.12 完全没有指纹的旧页码：只在内容流末尾兜底清理。
                #    如果还留着任何本工具指纹，就跳过——否则可能把刚保留的
                #    数字页眉当成旧页码删掉。
                if page_mode != "keep" and not _has_any_overlay_marker(new_ops):
                    new_ops = _strip_legacy_overlay_ops(
                        new_ops, (font_size, header_size),
                        cropbox_bottom=float(cb.bottom), page_height=vh,
                    )

                # ④ 文字层里可识别的原生页眉/页码（v1.17）。
                if header_mode == "clear" or page_mode == "clear":
                    fonts = _load_page_fonts(page)
                    items = _native_text_items(new_ops, fonts, cb, rot)

                    drop_header = None
                    if header_mode == "clear":
                        def drop_header(item):
                            return (
                                item["visual_y"] >= vh - NATIVE_BAND_PT
                                and _normalize_header_text(item["text"]) in native_header_texts
                            )

                    drop_page = None
                    if page_mode == "clear":
                        expected = {index}
                        label_number = _page_label_number(reader, index)
                        if label_number is not None:
                            expected.add(label_number)

                        def drop_page(item):
                            return (
                                item["visual_y"] <= NATIVE_BAND_PT
                                and _match_native_page_number(item["text"], expected, total)
                            )

                    new_ops, removed_headers, removed_pages = _remove_native_text_items(
                        new_ops, items, drop_header, drop_page,
                    )
                    native_header_count += removed_headers
                    native_page_count += removed_pages

                if new_ops != ops:
                    contents.operations = new_ops
                    page.replace_contents(ContentStream(ArrayObject([contents]), pdf=writer))

            # 只有 add 才画新的；keep / clear 都不画。
            text = format_page_text(start + index - 1, total, style) if page_mode == "add" else ""
            header_text = header if header_mode == "add" else ""
            if text != "" or header_text != "":
                overlay = make_overlay(
                    float(mb.width), float(mb.height), text,
                    font_name, font_size, x, y, rot, align,
                    header_text=header_text, header_font_name=header_font_name,
                    header_size=header_size, header_x=hx, header_y=hy,
                    header_align=header_align,
                )
                if overlay is not None:
                    # 叠加层的各个“框”必须和目标页一致：merge_page 默认按叠加层的 CropBox 裁剪，
                    # 框不一致会把页码裁掉或错位。
                    overlay.mediabox = mb
                    overlay.cropbox = cb
                    for box_name in ("/TrimBox", "/BleedBox", "/ArtBox"):
                        if box_name in page:
                            overlay[box_name] = page[box_name].get_object()

                    page.merge_page(overlay)

            # 合并会让 pypdf 把内容流解压成明文写回，不压回去输出体积会成倍膨胀
            # （实测：1.8MB 的扫描件能变成 5.9MB）。这里重新用 Flate 压缩。
            # 失败也不影响页码，只是文件偏大，所以忽略异常。
            try:
                page.compress_content_streams()
            except Exception:
                pass
        except Exception as exc:
            raise PdfNumberError(f"第 {index} 页处理失败：{exc}") from exc

    try:
        # 先写临时文件再改名，避免中途失败留下半个 PDF
        tmp = dst + ".part"
        with open(tmp, "wb") as fh:
            writer.write(fh)
        os.replace(tmp, dst)
    except PermissionError as exc:
        raise PdfNumberError("无法写入输出文件，目标可能正被其他程序打开") from exc
    except Exception as exc:
        raise PdfNumberError(f"保存失败：{exc}") from exc

    return total, font_warning, native_header_count, native_page_count


def write_status(path, message):
    """把中文结果写给 AHK 读（固定 UTF-8，和系统代码页无关）。"""
    if not path:
        return
    try:
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(message)
    except OSError:
        pass


def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="给 PDF 每页叠加页码（格式/对齐可调，原内容不改动）"
    )
    parser.add_argument("input", help="输入 PDF 文件")
    parser.add_argument("--out", required=True, help="输出 PDF 文件（可与输入相同则覆盖）")
    parser.add_argument("--start", type=int, default=1, help="第一页显示的页码，默认 1")
    parser.add_argument("--font-size", type=float, default=10.5, help="页码字号(pt)，默认 10.5")
    parser.add_argument("--margin-right", type=float, default=2.0,
                        help="水平边距(cm)：居右时距右边、居左时距左边、居中忽略；默认 2.0")
    parser.add_argument("--margin-bottom", type=float, default=1.2, help="距页面下边距(cm)，默认 1.2")
    parser.add_argument("--font-name", default="Helvetica", help="reportlab 内置字体名，默认 Helvetica")
    parser.add_argument("--format", dest="style", default="total",
                        choices=("plain", "dash", "total"),
                        help="页码格式：plain=1，dash=— 1 —，total=1 / n（默认）")
    parser.add_argument("--align", default="right",
                        choices=("left", "center", "right"),
                        help="页码水平位置：left/center/right，默认 right")
    parser.add_argument("--header", default="",
                        help="页眉文字；留空则用输入 PDF 的主文件名（不含扩展名）")
    parser.add_argument("--header-align", dest="header_align", default="center",
                        choices=("left", "center", "right"),
                        help="页眉对齐：left/center/right，默认 center")
    parser.add_argument("--header-size", type=float, default=0.0,
                        help="页眉字号(pt)，默认与 --font-size 相同")
    parser.add_argument("--header-margin-top", type=float, default=1.2,
                        help="页眉基线距页面上边(cm)，默认 1.2")
    # v1.15：两个“关掉”开关。两样都不画时仍会清理旧叠加层（等于撤销上一次加的页码/页眉）。
    # v1.17：每个项目三态。--keep-* 优先于 --no-*，但同时给两个视为参数错误，
    # 避免外壳拼错命令时悄悄按某一方执行。
    parser.add_argument("--no-page-numbers", action="store_true",
                        help="清空页码：删除本工具旧页码和文字层原生页码，不再添加")
    parser.add_argument("--keep-page-numbers", action="store_true",
                        help="不改变页码：原样保留已有页码，不删除也不添加")
    parser.add_argument("--no-header", action="store_true",
                        help="清空页眉：删除本工具旧页眉和文字层原生页眉，不再添加")
    parser.add_argument("--keep-header", action="store_true",
                        help="不改变页眉：原样保留已有页眉，不删除也不添加")
    parser.add_argument("--status", default="", help="把结果/错误写入这个文件（供 AHK 显示）")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(sys.argv[1:] if argv is None else argv)
    src = os.path.abspath(args.input)
    dst = os.path.abspath(args.out)

    if src == dst:
        write_status(args.status, "输入和输出是同一个文件，为安全起见请用不同文件名")
        return 2

    if args.keep_page_numbers and args.no_page_numbers:
        write_status(args.status, "--keep-page-numbers 和 --no-page-numbers 不能同时使用")
        return 2
    if args.keep_header and args.no_header:
        write_status(args.status, "--keep-header 和 --no-header 不能同时使用")
        return 2

    page_mode = "add"
    if args.keep_page_numbers:
        page_mode = "keep"
    elif args.no_page_numbers:
        page_mode = "clear"

    header_mode = "add"
    if args.keep_header:
        header_mode = "keep"
    elif args.no_header:
        header_mode = "clear"

    # 页眉文字：留空（含全是空白）就用输入 PDF 的主文件名。
    # 必须用 src（用户传进来的原始路径），不能用 dst——dst 是 AHK 的
    # “原名.pdf.pgnum.tmp”，取 basename 会变成 “原名.pdf.pgnum”。
    # keep / clear 时不传 --header；这里把 header 清空，避免误画。
    if header_mode == "add":
        header = args.header.strip()
        if not header:
            header = os.path.splitext(os.path.basename(src))[0]
    else:
        header = ""

    # 两项都“不改变”= 没有任何动作。仍然复制一份到 dst，保证外壳拿到输出文件；
    # 内容与原文件完全一致（正常 GUI 流程会在调用后端前就拦下这种组合）。
    if page_mode == "keep" and header_mode == "keep":
        try:
            shutil.copyfile(src, dst)
        except Exception as exc:
            write_status(args.status, f"复制文件失败：{exc}")
            return 2
        write_status(args.status, "未修改（页码和页眉都选择不改变）")
        return 0

    try:
        total, font_warning, native_header_count, native_page_count = add_page_numbers(
            src, dst,
            start=args.start,
            font_size=args.font_size,
            margin_right_cm=args.margin_right,
            margin_bottom_cm=args.margin_bottom,
            margin_top_cm=args.header_margin_top,
            font_name=args.font_name,
            style=args.style,
            align=args.align,
            header=header,
            header_align=args.header_align,
            header_size=args.header_size,
            page_mode=page_mode,
            header_mode=header_mode,
        )
    except PdfNumberError as exc:
        write_status(args.status, str(exc))
        return 2
    except Exception as exc:  # 兜底，避免只给 AHK 一个看不懂的 traceback
        write_status(args.status, f"未预期的错误：{type(exc).__name__}: {exc}")
        return 2

    # 状态文字要如实说明这次做了什么，外壳会原样显示在列表“状态”列和执行日志里。
    # 注意：AHK 用中文分号“；”判断“成功（有提示）”，所以正文里不要出现分号，
    # 只有 font_warning 是真正的提示，才以“；”开头接在后面。
    page_action = {"add": "添加页码", "clear": "清空页码", "keep": ""}[page_mode]
    header_action = {"add": "添加页眉", "clear": "清空页眉", "keep": ""}[header_mode]

    if page_action and header_action:
        message = f"已为 {total} 页{page_action}并{header_action}"
    elif page_action:
        message = f"已为 {total} 页{page_action}（页眉未改动）"
    elif header_action:
        message = f"已为 {total} 页{header_action}（页码未改动）"
    else:
        message = "未修改（页码和页眉都选择不改变）"

    details = []
    if native_page_count > 0:
        details.append(f"文字层原生页码 {native_page_count} 处")
    if native_header_count > 0:
        details.append(f"文字层原生页眉 {native_header_count} 处")
    if details:
        message += "（" + "、".join(details) + "）"
    write_status(args.status, message + font_warning)
    return 0


if __name__ == "__main__":
    sys.exit(main())
