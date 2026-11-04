# PDF 加页码工具 项目整体情况概览

> 交接文件。任何 AI 账号接手本项目前**先完整读完本文件**，复述当前状态与下一步计划，再按第 10 节「给下一个 AI 账号的接续指令」继续。
> 本文件在**每次开始改代码之前**更新（不是改完之后）。最后更新：2026-11-03 v1.18（已完成，待用户验收）——在 v1.17 上新建 `PDF 加页码工具 v1.18.ahk`：两个下拉框分组显示，页码下拉在「普通数字 / 格式数字 / 加总页码 / 不改变·清空」之间、页眉下拉在「对齐三项 / 不改变·清空」之间各画一条分隔线；Win32 下拉框没有原生分隔线，改成 owner-draw 自绘（创建时 `+0x210`，处理 `WM_DRAWITEM`），占位项 `PDF_SEP` 只画 1px 横线不画文字，选中时由 `PdfSepGuard()` 跳到相邻正常项，不会写进注册表；顺手删掉页眉列表里遗留的 `"-"`。只改 AHK 外壳，`pdf_page_number.py` 与 `build\backend.zip`（19,123,458 B / 33 项）未动，只重编译 EXE（20,505,088 B）。已实测：`/validate` EXIT=0；EXE 静默跑 11 页 PDF 副本 exit=0（原地改写、备份、日志、`backend-1.18\backend.ok`=1.18）；GUI 16 项检查 0 失败（分隔线位置/截图灰线/键盘与鼠标两条守卫路径）。当前版本 **v1.18**。

## 1. 项目目标与范围

`w:\eyes\Desktop\公文` 仓库里有两个互不依赖的工具，本文件只覆盖**第二个**：

- 主项目：`6.8.19.ahk` 党政机关公文自动排版引擎，另有 `程序整体说明.md` / `AGENTS.md` 维护，**不要在本次任务里改动其逻辑**。
- 本项目：**PDF 加页码小工具**。给已有 PDF 每页叠加页码，原地写回原文件，写前自动备份；不改变原 PDF 内容流（不重排版面，扫描件也能加）。
  - 范围：AHK v2 外壳（窗口/拖放/右键菜单/备份/日志/提示）+ Python 后端（pypdf + reportlab）。
  - 不做：OCR、拆分合并、压缩、加密解密、页码之外的排版。

## 2. 当前整体进度

| 版本 | 状态 | 内容 |
| --- | --- | --- |
| v1.4 | **已完成并验证** | 9 种页码格式（3 格式 × 左/中/右）、单一水平边距、右键菜单改名、列表右键菜单（打开文件/打开目录/恢复上一版本）、新图标编译 |
| v1.5 | **已完成（2026-10-30 13:50）** | 参数区 4 列等距 + 标签右对齐、页码格式下拉通栏、纵向重新分配、窗口 700×705 → 700×685；EXE 已重编译，`backend-1.5` 已解压验证 |
| v1.6 | **已完成（2026-10-30 14:07）** | 除「拖入含子文件夹询问递归」「恢复上一个版本覆盖确认」「选择备份文件夹」三处必须由用户选择的地方外，其余模态框全部改为右下角浮动提示（`ShowPdfToast`）＋状态栏；删掉“结果已同时在屏幕右下角提示”元描述；EXE 已重编译，`backend-1.6` 已验证 |
| v1.7 | **已完成（2026-11-02）** | P0 项：`#SingleInstance Force`→`Off`、`ProcessOne` 缺文件检查、`ReplaceOriginalPdf` 的 `FileMove` 重试、恢复前快照 `_before-restore_<时间戳>`、`NormalizePdfParams` 参数归一；后端与 zip 未变，只重编译 EXE |
| v1.8 | **已完成（2026-11-02）** | 备份目录行「选择…」按钮加高（25→40），新增「打开」按钮（`OpenPdfBakFolder`）；只重编译 EXE |
| v1.9 | **已完成（2026-11-02）** | 页码参数改动即保存（控件 `Change` → `SavePdfParamNow`，启动由 `LoadPdfSettings` 回填）；只重编译 EXE |
| v1.10 | **已完成（2026-11-02）** | 后端新增旧页码剥离：`_strip_trailing_page_number_ops()` 先去掉尾部旧页码绘制再叠加新页码；**改了 `pdf_page_number.py`，已重打 `build\backend.zip` 并重编译 EXE** |
| v1.11 | **已完成（2026-11-02）** | 修复文件头修改记录乱码、补齐 v1.7–v1.10 记录、修正备份目录行「打开」按钮文案（原为 `å¼æ¬`）；仅注释与界面文案，重编译 EXE（20,473,856 B）；静默跑副本 exit=0、`backend-1.11` 解压正常 |
| v1.12 | **已完成（2026-11-02）** | 修复 `SetPdfControlsLocked()` 漏 `uiPdfBakOpen` 的 `global` 声明（任何批量处理必崩）；`AddPdfFile()` 补重复文件计数；`Iconi` → `Icon!`；源码态右键菜单改用 `PDF 加页码工具.exe`；只改 AHK 外壳，重编译 EXE（20,473,856 B） |
| v1.13 | **已完成，待用户验收（2026-11-02 15:07）** | 新增「页眉」：文字留空则用 PDF 主文件名，三选对齐（默认居左），可调「距上边(cm)」（默认 1.2），字号与页码共用；后端一次 overlay 画页眉+页码并注册中文字体；旧叠加层清理改为「标记块优先 + 尾部页码形状兜底」，避免按字号全流扫描误删 10.5pt 正文；程序文件名带版本号（`v1.13`）；后端已重打 zip、EXE 已重编译，`/validate` 与源码/EXE 静默跑副本均 exit=0；下一步只剩 GUI/右键菜单人工验收 |
| v1.15 | **已完成，待用户验收（2026-11-03 09:31）** | 界面微调 + 两个开关：「选择…」与左侧输入框统一 `h28` 同一 `y`（垂直对齐），「打开」改成 28×28 正方形图标按钮（`BS_ICON` + `BM_SETIMAGE` + `SHGetFileInfoW` 系统文件夹图标，悬停提示用 `WM_MOUSEMOVE` + `ToolTip()` 自己实现——AHK v2 的 `Gui.Button` 没有 `ToolTip()` 方法）；`PDF_FORMAT_LIST` 首项新增「无」= 不加页码，起始页码/水平边距/距下边置灰；页眉组框标题右侧新增「添加页眉」复选框（`AddHeader`，默认 1），取消勾选 = 只加页码；后端新增 `--no-page-numbers` / `--no-header`，状态文案分三种；两样都关时 GUI 直接提示返回、静默模式记失败且不备份不改原文件。**改了 `pdf_page_number.py` → 已重打 `build\backend.zip`（19,551,183 B）并重编译 EXE（20,928,512 B）** |

| v1.16 | **已完成，待用户验收（2026-11-03 10:44）** | 以 **v1.14** 为基线（不包含 v1.15 的页码「无」/`--no-page-numbers`/`AddHeader` 复选框）：① 备份目录行 `Edit`/「选择…」/「打开」统一 `h28`、同一 `y`，「打开」改 28×28 系统文件夹图标按钮（`BS_ICON` + `BM_SETIMAGE` + `SHGetFileInfoW`），悬停提示用 `WM_MOUSEMOVE` + `ToolTip()`；② 页眉下拉首项新增「不添加页眉」（`PDF_HEADER_NONE`），标签「对齐」→「页眉」、组框→「页眉（可选）」，选它时页眉文字/距上边置灰、`ProcessOne()` 改传 `--no-header`，默认值即「不添加页眉」；只改 AHK 外壳（后端复用已含 `--no-header` 的 `build\backend.zip`），EXE 20,927,488 B；`/validate`、源码/EXE 静默（有页眉/无页眉两条）均通过 |
| v1.17 | **已完成，待用户验收（2026-11-03 13:01）** | 在 v1.16 上做三态：页码下拉末尾加「不改变页码」「清空页码」，页眉下拉改为「不改变页眉 / 居左 / 居中 / 居右 / 清空页眉」（默认两项都是「不改变」）；`ApplyPdfHeaderState()` 扩成 `ApplyPdfUiState()` 统一置灰；`ProcessOne()` 按选择拼 `--keep-page-numbers` / `--no-page-numbers` / `--keep-header` / `--no-header`；两项都「不改变」时 GUI 提示返回、静默/右键在备份前记失败。后端 `pdf_page_number.py` 新增 `--keep-*`、页码/页眉独立指纹块（1.33 / 1.32，旧合并层 1.31 按视觉 y 拆分）、文字层原生页码/页眉识别（页脚/页顶 4cm 带）；**改了 py → 已重打 `build\backend.zip`（19,123,458 B / 33 项）并重编译 EXE（20,501,504 B）**；`/validate` EXIT=0，后端 78 项检查 0 失败，源码/EXE 静默 7 组组合通过 |
| v1.18 | **已完成，待用户验收（2026-11-03 13:54）** | 下拉框分组：页码下拉在「普通数字 / 格式数字 / 加总页码 / 三态项」之间、页眉下拉在「对齐三项 / 三态项」之间加分隔线。Win32 下拉框没有原生分隔线，改用 owner-draw（`+0x210` = `CBS_OWNERDRAWFIXED|CBS_HASSTRINGS` + `WM_DRAWITEM` 自绘），占位项 `PDF_SEP` 只画 1px 横线不画文字；`PdfSepGuard()` 在选中分隔条时按方向跳到相邻正常项（鼠标点击退回原值），所以它永远不会进注册表。`PDF_FORMAT_LIST` / `PDF_HEADER_ALIGN_LIST` 保持纯数据，另建 `*_ITEMS` 给界面用；删掉页眉列表遗留的 `"-"`。只改 AHK 外壳（py 与 zip 不动，zip 仍 19,123,458 B / 33 项），重编译 EXE（20,505,088 B）；`/validate` EXIT=0、EXE 静默跑 11 页副本 exit=0、GUI 16 项检查 0 失败 |
| v1.14 | **已完成（2026-11-03 08:52）** | 修复编译版在沙箱/网盘重定向目录里“后端没有生成输出文件”（根因：Sandboxie `OpenFilePath=w:\Eyes\Desktop\公文` 导致冻结后端写入落到另一层视图，文件其实写出来了但外壳看不到）；`ProcessOne()` 改为后端只写 `%TEMP%`、外壳复制回原目录并校验大小后原子改名；失败原因补退出码/后端状态/临时文件路径；只改 AHK 外壳，重编译 EXE（20,925,440 B）；`/validate` 与三条运行路径均验证通过 |

v1.6 只改提示出口；v1.7–v1.9、v1.11–v1.12 只改 AHK 外壳（后端 `pdf_page_number.py` 与 `build\backend.zip` 未变，只需重编译 EXE）；v1.10 与 v1.13 动了后端，必须重打 zip。v1.14 只改 AHK 外壳（后端与 zip 未变），修的是沙箱目录写入失败；v1.16 同样只改 AHK 外壳（后端与 zip 未变），只重编译 EXE（20,927,488 B）；v1.17 动了后端，已重打 zip（19,123,458 B）并重编译 EXE（20,501,504 B）；v1.18 只改 AHK 外壳（下拉框分组分隔线），不动后端与 zip；包体积/杀软优化仍见第 5.6 节，等用户拍板。

### 2.1 v1.7 实施计划（AI 代码审查归档，2026-10-30，待批准 → 已实施）

- 状态：AI 已完成 `PDF 加页码工具.ahk` 代码审查并把发现归档到第 8.1 节；用户已口头确认「要的」后明确批准「可以编写代码了」（2026-11-02）；**v1.7 代码已全部落地并编译通过**（`/validate` EXIT=0，EXE 重编译，后端与打包产物未变，未重打 zip）；等待用户测试验收。
- 范围：仅修 `PDF 加页码工具.ahk` 的 P0 项，后端 `pdf_page_number.py` 与打包产物不变。
- 落地摘要：单实例 `Off`；`ProcessOne` 缺文件检查；`IsPdfWritable` 前置 `FileExist`；`ReplaceOriginalPdf` FileMove 重试；恢复前快照；新增 `NormalizePdfParams`。
- 未进入 v1.7 的 P1/P2 项先记入候选清单（第 8.1 节末尾），由用户后续决定排期。

### 2.2 v1.13 实施记录（2026-11-02，已完成）

- 状态：**已完成（2026-11-02 15:07）**。用户已确认 4 个细节：①新增「距上边(cm)」输入框（不写死）；②页眉默认居左；③旧的无版本号 `.ahk`/`.exe` **保留**；④页眉与页码共用字号。
- 目标文件：
  - 新建 `PDF 加页码工具 v1.13.ahk`（从 v1.12 复制后修改，**不要直接改 v1.12**）。
  - `pdf_page_number.py`（后端，直接改；改完必须重打 `build\backend.zip`）。
  - 编译输出 `PDF 加页码工具 v1.13.exe`（Ahk2Exe 的 `/out` 指定文件名）。
- 后端改动 `pdf_page_number.py`：
  1. 新增 `register_cjk_font()`：按 `C:\Windows\Fonts\msyh.ttc`（subfontIndex=0）→ `simsun.ttc` → `simhei.ttf` → `None`（回退 Helvetica）注册 reportlab TTFont，结果缓存；返回字体名与是否成功。
  2. `make_overlay()` 增加页眉参数，在同一张 overlay 上画「页眉（顶部）」+「页码（底部）」，仍只 `merge_page` 一次；页眉走中文字体，页码数字也用同一字体（避免一页两种字形）。
  3. 新增命令行参数：`--header`（空串=用输入 PDF 的主文件名）、`--header-align left|center|right`（默认 center，AHK 传 left）、`--header-margin-top`（cm，默认 1.2）、`--header-size`（默认取 `--font-size`）。
  4. 旧叠加层清理重写：新 overlay 在内容流里画一个**不可见标记文本**（render mode 3，内容 `DocFmtPN`），清理时按标记定位、整块删除（先删最近的、循环到没有标记为止）；没有标记的旧版（v1.10–v1.12）叠加层再用「尾部、页码形状、同字号、最多 4 块」的兜底逻辑清一次。**绝不按字号扫描整个内容流**（正文五号=10.5pt 会被误删，这是 v1.10 起的潜在事故）。
  5. `--status` 文案：有页眉时写「已为 N 页添加页码和页眉」。
- AHK 改动（`PDF 加页码工具 v1.13.ahk`）：
  1. `PDF_TOOL_VERSION := "1.13"`；文件头「修改记录」加 v1.13。
  2. 新增常量 `PDF_HEADER_TEXT`（默认空）、`PDF_HEADER_ALIGN`（默认「居左」）、`PDF_HEADER_MARGIN_TOP`（默认 1.2）、`PDF_HEADER_ALIGN_LIST := ["居左","居中","居右"]`。
  3. 界面新增「页眉」组框 `x20 y225 w660 h105`：页眉文字 `Edit x125 y246 w300`（Limit 100）、对齐 `DropDownList x483 y246 w122`、距上边 `Edit x125 y278 w55`；灰色说明一行。后续下移：`选项` y340、`任务列表` y505、按钮行 y730，窗口 `700×800`。
  4. 新增 `global` 声明、`SetPdfControlsLocked()` 列表、`RunGuiProcessing()` 校验（距上边 0–20）与写回、`LoadPdfSettings()`、`SavePdfParamNow()`；新增注册表值 `HeaderText`（**允许空串，必须用直接 RegRead 的 `GetPdfRawStr()`，不能用 `GetPdfStr()`**）、`HeaderAlign`（白名单）、`HeaderMarginTop`（0–20）。
  5. `ProcessOne()` 拼 `--header`（空串也要显式传）、`--header-align`、`--header-margin-top`。
  6. 源码态查找同目录 EXE 的硬编码 `"PDF 加页码工具.exe"` 改为 `RegExReplace(A_ScriptName, "i)\.ahk$", ".exe")`。
- 文档改动：`程序整体说明.md` §13、本文件、`AGENTS.md`（PDF 工具「文件名不能改」的旧规则改为「文件名带版本号」＋右键菜单需重勾一次）。
- 验证计划：`/validate`；
- 已知风险：页眉覆盖在原内容上方，原文顶部 1.2cm 内有内容会重叠；页眉含 emoji/生僻字可能缺字；旧无版本号 EXE 保留会让右键菜单继续指向旧路径，需用户在新窗口里取消再勾选一次。
- 实际结果（2026-11-02 15:07）：`/validate` EXIT=0；后端重打 `build\backend.zip` 19,551,127 B；EXE 21,303,296 B；源码态与 EXE 静默跑 11 页副本均 exit=0，日志写「已为 11 页添加页码和页眉」；页眉基线 y≈807.9、页码 y≈34.02（A4/1.2cm/10.5pt，居左 x≈56.69）。**尚未做 GUI 人工验收与右键菜单重勾。**

### 2.3 v1.14 实施记录（2026-11-03，已完成，待用户验收）

- 触发：用户报告 `PDF 加页码工具 v1.13.exe` 加页码失败，同一份 PDF 用 `PDF 加页码工具 v1.13.ahk`（源码态）却成功；执行日志里全是“后端没有生成输出文件”。
- 根因（实测，非推测）：`C:\Windows\Sandboxie.ini` 的 `[DefaultBox]` 只有两条放行规则——`OpenFilePath=w:\Eyes\Desktop\公文`、`OpenFilePath=c:\python`。命中该前缀规则的目录（`公文`/`公文2`/`公文测试`）里，**冻结后的** `pdf_page_number.exe` 写出的 `.pgnum.tmp` 会落到另一层文件视图：文件确实存在（工作区文件工具可见，能取到完整 2.2 MB 内容），但外壳 AHK 与 PowerShell 的 `Test-Path`/`FileExist()` 都看不到，于是误判失败。
- 对照实验：同参数用 `C:\Python\python.exe pdf_page_number.py` 写同一目录 → 成功（`c:\python` 同样在放行列表里，所以源码态一直正常）；把冻结 exe 复制到 `%TEMP%` 再写该目录 → 仍失败（排除 exe 路径/工作目录）；写 `我的公文`、`ABC公文`、`GongWen`、`测试3`、`w:\eyes\Desktop\` 根、`w:\eyes\Documents`、`%TEMP%` → 全部正常（与“以 `公文` 开头”的前缀匹配吻合）；`cmd`/PowerShell/编译版 AHK 的 `FileAppend`/`FileCopy`/`FileMove` 写该目录也全部成功。
- 修复（`ProcessOne()` ③④ 两步）：③ 后端输出改为 `%TEMP%\pdf_pagenum_<PID>_<TickCount>_<序号>.pdf`（带 PID，因为 `#SingleInstance Off` 会有多实例并发）；④ `FileCopy` 到原目录 `原名.pdf.pgnum.tmp` → `FileGetSize` 大小校验（0 或不等即判失败）→ `ReplaceOriginalPdf()` 原子改名覆盖。每一步失败都只删中间/临时文件，原文件一字节不动；成功路径最后删掉 `%TEMP%` 里的中转文件。
- 顺手改：退出码 0 但拿不到输出时，失败原因带上“退出码 + 后端写入的状态文本 + 临时文件路径”，不再只有一句“后端没有生成输出文件”。
- 只动 AHK 外壳：`pdf_page_number.py` 与 `build\backend.zip`（19,551,127 B）未变，跳过 PyInstaller 直接重编译 EXE。
- 验证结果：`/validate` EXIT=0；新 EXE 20,925,440 B；首跑解压出 `backend-1.14\backend.ok`＝`1.14`；源码态 `--quiet`、EXE `--quiet`、EXE 真实窗口（`WM_DROPFILES` 拖放 + 真实鼠标点击）三条路径都在 `w:\eyes\Desktop\公文` 内跑成功（2,202,851 → ≈2,204,0xx B，日志“已为 11 页添加页码和页眉”，无 `.pgnum.tmp` 残留，`pypdf` 抽文字可见页眉文件名与 `— 1 —`/`— 11 —`）。**待用户 GUI 验收 + 右键菜单重勾一次。**

### 2.4 v1.15 实施记录（2026-11-03，已完成，待用户验收）

用户反馈（原话）：v1.14 功能没问题，只要改界面——备份目录行的「选择…」要和左边输入框一样高（或略高）并垂直对齐；
「打开」不要文字、要用图标、正方形、高度与「选择…」一致；另外要能自己决定是否加页眉，页码则直接在格式里选「无」。

落点与边界：

| # | 改动 | 文件 | 边界/注意 |
| --- | --- | --- | --- |
| 1 | 备份目录行：`Edit` `h28 y393`；`选择…` `x545 y393 w80 h28`；`打开` `x635 y393 w28 h28` | `.ahk` | 三个控件同一 `y`、同一 `h`，否则又会像 v1.8 那样“按钮比输入框高一截” |
| 2 | 「打开」改图标按钮：`+0x40`（`BS_ICON`）+ `BM_SETIMAGE`（`0x00F7`, `IMAGE_ICON=1`），图标用 `SHGetFileInfoW` 取系统文件夹小图标（16px，28px 按钮内不会被裁） | `.ahk` | 取图标失败时按钮仍可点，只是看不到图；**AHK v2 的 `Gui.Button` 没有 `ToolTip()` 方法**（实测抛 “has no method named ToolTip”），改用 `WM_MOUSEMOVE` + `ToolTip()` 自己实现悬停提示 |
| 3 | `PDF_FORMAT_LIST` 首项加「无」 | `.ahk` | `PdfFormatArgs()` 不用为「无」加分支：「无」不含任何关键词，自然落到 `plain` + `center`，而 `ProcessOne()` 会按「无」追加 `--no-page-numbers`（样式参数后端不读）；`LoadPdfSettings()` 的合法性校验天然兼容。默认值仍是「加总页码 1 / n 居右」 |
| 4 | 页眉开关：「添加页眉」复选框 `PDF_ADD_HEADER`（注册表 `AddHeader`，默认 1） | `.ahk` | 与 `PDF_CREATE_BAK` 同款：`GetPdfNum("AddHeader", 1) ? 1 : 0`；勾选即存 |
| 5 | 联动置灰：选「无」→ 起始页码/水平边距/距下边禁用；不勾「添加页眉」→ 页眉文字/对齐/距上边禁用 | `.ahk` | 不能直接写 `Enabled := false`，要先看批量锁 `PdfUiLocked`，解锁时再按开关重算 |
| 6 | 传参：`--no-page-numbers` / `--no-header` | `.ahk` + `.py` | 页眉留空在后端是“用文件名”，所以“不加页眉”不能用空串表达，必须是独立开关 |
| 7 | 后端状态文案分三种：页码+页眉 / 只有页码 / 只有页眉 | `.py` | AHK 的「成功（有提示）」判断靠中文分号 `；`，新文案里不要多加分号 |
| 8 | 两个都关（格式=无 且 不勾页眉） | `.ahk` | 没有可添加的内容：GUI 点「开始」前直接提示并返回；静默模式每个文件记一条失败，**不备份、不改原文件** |

验证计划：`/validate` → 截图核对对齐 → 源码态与 EXE 各跑三种组合（无页码+页眉 / 页码+无页眉 / 两者都要）→ 确认状态文字、旧叠加层清理、无 `.pgnum.tmp` 残留 → 更新 `程序整体说明.md` 第 13 节与 `AGENTS.md` → 启动窗口，用户重勾右键菜单。

实施结果（2026-11-03 09:31）：

- 后端（`pdf_page_number.py`，36,651 B，无 BOM + LF，`py_compile` 通过）四种组合直跑全部 rc=0、11 页：both（页眉 y≈807.9 + 页码 y≈34.0）/ `--no-header`（只剩页码）/ `--no-page-numbers`（只剩页眉）/ 两个都给（overlay 为空，只清理旧层）。状态文本依次为「已为 11 页添加页码和页眉」「已为 11 页添加页码」「已为 11 页添加页眉」「已处理 11 页（未添加页码和页眉）」。
- 界面（`/validate` EXIT=0，窗口 700×800 客户区）：实测控件几何 `Edit7(备份目录) x=105 y=393 w=430 h=28`、`Button6(选择…) x=545 y=393 w=80 h=28`、`Button7(打开) x=635 y=393 w=28 h=28`，三者同一 `y` 同一高度；`BM_GETIMAGE` 取到非零图标句柄（系统文件夹图标）。
- 联动置灰：页码格式选「无」→ 起始页码/水平边距/距下边 `IsWindowEnabled=0`，字号仍为 1；取消「添加页眉」→ 页眉文字/对齐/距上边全为 0；勾回/换回后全部恢复为 1。
- EXE（20,928,512 B）静默跑四份副本：A 无页码+页眉 rc=0（只出现页眉）、B 页码+无页眉 rc=0（只出现页码）、C 两样都关 rc=1 且 md5 未变、备份数不变、D 两者都要 rc=0；全程无 `.pgnum.tmp` 残留，日志四行结果与后端文案一致，首跑解压出 `backend-1.15\backend.ok`＝`1.15`。
- 编码：`.ahk` UTF-8 带 BOM + LF（106,878 B），`.py` 无 BOM + LF，`build\backend.zip` 根下为 `pdf_page_number.exe` + `_internal\`。
- 剩余：用户在窗口里重勾一次右键菜单（`RepairPdfContextMenu()` 只在旧 EXE 路径失效时才切换，v1.14 的 EXE 还在）。

### 2.5 v1.16 实施记录（2026-11-03，已完成，待用户验收）

用户确认：以 **v1.14** 为基线新建 v1.16（v1.15 保留不动），只做界面调整，不引入 v1.15 的页码「无」/`--no-page-numbers`/`AddHeader` 复选框。

- 备份目录行：`Edit`、`选择…`、`打开` 统一 `h28`、同一 `y`（`x105/x545/x635 y393`）；「打开」改 28×28 图标按钮（`BS_ICON` + `BM_SETIMAGE` + `SHGetFileInfoW` 取系统文件夹图标），不显示文字，并加悬停提示「打开备份目录」。
- 页眉下拉：`PDF_HEADER_ALIGN_LIST` 首项加「不添加页眉」，默认值「不添加页眉」（首次安装/清空注册表后只加页码；注册表已存的 HeaderAlign 不覆盖）；标签「对齐」改「页眉」，组框标题改「页眉（可选）」。
- 选「不添加页眉」时：页眉文字、距上边置灰；`RunGuiProcessing()` 跳过页眉相关校验；`ProcessOne()` 传 `--no-header` 而不是 `--header*`。批量处理结束解锁后要重新应用置灰状态。
- 后端：复用当前 `build\backend.zip`（`pdf_page_number.py` 已含 `--no-header`），不改 py，只重编译 AHK。
- 实际结果（2026-11-03 10:44，默认值定稿为「不添加页眉」后复测）：`PDF 加页码工具 v1.16.ahk`（UTF-8 BOM + LF），`/validate` exit=0；EXE 重编译 20,927,488 B（10:43:34）；删掉注册表 `HeaderAlign` 后跑副本 → 只剩页码，日志「已为 11 页添加页码」；把 `HeaderAlign` 设为「居左」→ 页眉取文件名 + 页码，日志「已为 11 页添加页码和页眉」，exit=0；首跑解压 `backend-1.16\backend.ok`=1.16。
- 剩余：用户 GUI 验收（按钮高度/图标/悬停提示/下拉置灰）＋ 在新窗口把右键菜单取消再勾选一次（`RepairPdfContextMenu()` 只在旧 EXE 路径失效时切换，v1.15 的 EXE 还在）。
- 注意：`v1.15` 保留未动；它的页码「无」/`AddHeader` 复选框功能在 v1.16 里没有，若用户要“只加页眉”仍用 v1.15。


### 2.6 v1.17 实施记录（2026-11-03，已完成，待用户验收）

用户确认：在 v1.16 基础上新建 v1.17，页码和页眉都改成三态；「清空」除删除本工具叠加层外，还要删除**文字层**里可识别的原生页眉/页码（扫描图片里印着的内容不处理）。

**开工前状态（2026-11-03 11:05）**

- 后端 `pdf_page_number.py` 已完成 v1.17 逻辑（60,655 B，无 BOM + LF）：`page_mode/header_mode` 三态、`--keep-page-numbers` / `--keep-header`、页眉/页码独立指纹块、v1.13–v1.16 合并层按位置拆分、文字层原生页眉/页码识别（顶部/底部 4cm 带）。**尚未重打 `build\backend.zip`**（当前仍是 9:21 的旧包，不含 `--keep-*`）。
- `PDF 加页码工具 v1.17.ahk` 目前只是 v1.16 的副本：`PDF_TOOL_VERSION` 仍是 `"1.16"`，只把页眉默认值改成「不添加页眉」；三态界面未写，也没有 `v1.17.exe`。

**本轮实施计划**

- 页码下拉 `PDF_FORMAT_LIST` 末尾新增「不改变页码」「清空页码」；默认仍是「加总页码 1 / n 居右」。
- 页眉下拉 `PDF_HEADER_ALIGN_LIST`：首项「不添加页眉」改为「不改变页眉」，末尾新增「清空页眉」；默认「不改变页眉」。
- 传参：`ProcessOne()` 按选择拼 `--keep-page-numbers` / `--no-page-numbers` / `--keep-header` / `--no-header`；只有“添加”才传 `--header*` 和页码格式参数。
- 置灰：`ApplyPdfHeaderState()` 扩成 `ApplyPdfUiState()`，同时管页码三项（起始页码/水平边距/距下边）、字号、页眉两项；仍先看批量锁 `PdfUiLocked`。
- 校验：`RunGuiProcessing()` 只校验本次真正会用到的参数；页码和页眉都选「不改变」时 GUI 直接提示返回，静默/右键模式在 `ProcessOne()` 备份前记失败、不改原文件。
- 版本：`PDF_TOOL_VERSION := "1.17"`、文件头加 v1.17 修改记录、窗口标题随版本。
- 打包顺序：重打 `build\backend.zip`（py 已变）→ Ahk2Exe 编译 `PDF 加页码工具 v1.17.exe`。
- 验证：`/validate`；源码态与 EXE 各跑组合（清页码+清页眉 / 清页码+加页眉 / 加页码+清页眉 / 保持+添加 / 两项都保持）；核对原地改写、备份、日志、无 `.pgnum.tmp` 残留、`backend-1.17\backend.ok`=`1.17`、文字层原生页眉/页码被清、旧叠加层按类保留。
- 已知限制：扫描图片里印着的页眉/页码无法删除；文字层识别是保守启发式，可能漏删，但不会按字号扫全文、不会为了删页眉去删正文。

**实际结果（2026-11-03 13:01）**

- 外壳：`PDF 加页码工具 v1.17.ahk`（UTF-8 BOM + LF，111,562 B / 2,408 行）：`PDF_TOOL_VERSION="1.17"`；`PDF_FORMAT_LIST` 11 项、`PDF_HEADER_ALIGN_LIST` 5 项；新增 `PDF_PAGE_KEEP` / `PDF_PAGE_CLEAR` / `PDF_HEADER_KEEP` / `PDF_HEADER_CLEAR`；`OnPdfFormatChange()` + `ApplyPdfUiState()`；`RunGuiProcessing()` 只校验本次用到的参数、两项都「不改变」直接提示返回；`ProcessOne()` 在备份前拦 no-op、按三态拼开关。`/validate` EXIT=0。
- 后端回归（源码态，`W:\Eyes\Temp\opencode\run_v117_backend_tests.py`）：78 项检查 0 失败。覆盖原生文字层清空（页码/页眉删除、正文保留）、真实 11 页 PDF 的 v1.16 合并层拆分、添加/清空/保留各种组合、keep+keep 字节一致、独立指纹 1.32/1.33。
- 打包：PyInstaller 重新生成 `build\dist\pdf_page_number\`（exe 6,227,290 B）；`build\backend.zip` 19,123,458 B / 33 项，根下 `pdf_page_number.exe` + `_internal/`，无反斜杠条目；`tar.exe -xf` 解压 EXIT=0。
- 冻结后端：keep+keep 输出与输入 MD5 一致、状态「未修改（页码和页眉都选择不改变）」；clear+clear 与源码态输出 MD5 一致、状态「已为 11 页清空页码并清空页眉」。
- EXE：`PDF 加页码工具 v1.17.exe` 20,501,504 B（2026-11-03 12:54）；首跑解压 `%LOCALAPPDATA%\PdfPageNumberTool\backend-1.17\`，`backend.ok` 内容 `1.17`（UTF-8 BOM + 文本）。
- 外壳端到端（`W:\Eyes\Temp\opencode\run_v117_shell_e2e.py`，7 组、35 项检查）：源码态清/清、EXE keep/keep（exit=1、不备份、字节不变、日志记失败）、EXE 加/清、清/加、keep/加、加/keep、源码态加/keep 全部符合预期；每组都确认原地改写、备份是原件、无 `.pgnum.tmp` 残留、日志追加。
- 剩余：用户 GUI 验收（三态置灰、默认两项「不改变」、清空/添加组合）＋ 在新窗口把右键菜单取消再勾选一次（`RepairPdfContextMenu()` 只在旧 EXE 路径失效时切换，v1.16 的 EXE 还在）。

### 2.7 v1.18 实施记录（2026-11-03，已完成，待用户验收）

用户要求：下拉框里各选项分组，组与组之间加一条分隔线。

**技术前提（已在本机实测）**：Win32 的下拉框没有原生分隔线，只能 owner-draw 自绘。验证脚本 `W:\Eyes\Temp\opencode\test_sep_combo.ahk`：

- `+0x210` 生效，实际窗口样式 `0x50010213`（含 `CBS_OWNERDRAWFIXED` 0x10、`CBS_HASSTRINGS` 0x200、`CBS_DROPDOWNLIST` 0x3）；
- `DRAWITEMSTRUCT` 64 位偏移正确（0 CtlType / 4 CtlID / 8 itemID / 12 itemAction / 16 itemState / 24 hwndItem / 32 hDC / 40 RECT / 56 itemData），
  12 个条目逐条收到 `WM_DRAWITEM`；
- 截图取像素核对：3 条分隔线各 286 px 长（颜色 `0xC8C8C8`），线之间是深色文字像素 —— 横线与文字都正常绘制；
- 选中分隔条时按方向跳到相邻正常项，上下键都能穿过（不会被卡住）。

**计划改动（只动 AHK 外壳）**：

1. 新建 `PDF 加页码工具 v1.18.ahk`（从 v1.17 复制），`PDF_TOOL_VERSION` → `"1.18"`，文件头补 v1.18 修改记录。
2. 新增 `PDF_SEP := Chr(0x2500)` 占位标记；`PdfAddSeps(src, afterList)` 把纯数据列表按位置插线；
   `PDF_FORMAT_ITEMS := PdfAddSeps(PDF_FORMAT_LIST, [3, 6, 9])`、`PDF_HEADER_ITEMS := PdfAddSeps(PDF_HEADER_ALIGN_LIST, [3])`。
3. 两个下拉框创建加 `+0x210`，条目改用 `*_ITEMS`；`Choose` 的索引也按 `*_ITEMS` 找。
4. `ShowPdfGui()` 里注册 `OnMessage(0x002B, PdfSepDraw)`（`WM_DRAWITEM`）。
5. `PdfSepDraw()`：分隔条铺背景 + 中线画 1px 灰线；普通项自绘高亮/普通背景 + 文字（禁用时画 `COLOR_GRAYTEXT`）。
6. `PdfItemText()`（`CB_GETLBTEXT`）、`PdfSepGuard()`（选中分隔条时按方向跳、鼠标点击退回原值）。
7. `OnPdfFormatChange()` / `OnPdfHeaderAlignChange()` 开头先过 `PdfSepGuard()`；同时删掉 v1.17 里 `PDF_HEADER_ALIGN_LIST` 的那个 `"-"`
   （它是个可正常选中的普通项，选中会被当成“添加页眉、居中”）。
8. 验证：`/validate`；启动 EXE 后用 `CB_SHOWDROPDOWN` + 截图像素核对分隔线数量与位置；
   用 `CB_SETCURSEL` + 伪造 `WM_COMMAND(CBN_SELCHANGE)` 验证点分隔条不会写进注册表；
   再跑一份 PDF 副本的 `--quiet` 确认业务没被界面改动影响。
9. 只改 `.ahk` → **不重打 zip**，直接 Ahk2Exe 编译；升版本后要重新勾一次右键菜单。

**实际结果（2026-11-03 13:54）**：

- 外壳：`PDF 加页码工具 v1.18.ahk`（UTF-8 带 BOM + LF，120,835 B / 2,596 行）：`PDF_TOOL_VERSION="1.18"`；`PDF_FORMAT_LIST` 11 项、`PDF_HEADER_ALIGN_LIST` 5 项（删掉遗留的 `"-"`）；新增 `PDF_SEP`、`PdfAddSeps()`、`PDF_FORMAT_ITEMS`（14 项）/ `PDF_HEADER_ITEMS`（6 项）；两个下拉加 `+0x210`；`ShowPdfGui()` 里 `OnMessage(0x002B, PdfSepDraw)`；新增 `PdfSepDraw()` / `PdfItemText()` / `PdfSepGuard()`；两个 Change 处理先过守卫。`/validate` EXIT=0。
- EXE：`PDF 加页码工具 v1.18.exe` 20,505,088 B（2026-11-03 13:51）；只改 AHK 外壳，**没有重打 `build\backend.zip`**（仍是 19,123,458 B / 33 项）。
- 后端首跑：EXE 静默跑 11 页 PDF 副本 exit=0，原文件被原地改写（3,647,786 → 3,649,437 B），备份 `v118_quiet_test_20261103_135401.pdf`、日志追加一行；首跑解压 `%LOCALAPPDATA%\PdfPageNumberTool\backend-1.18\`，`backend.ok` 内容 `1.18`。
- GUI 验证（`W:\Eyes\Temp\opencode\test_v118_gui.py`，16 项检查 0 失败）：窗口标题 `PDF 加页码工具 V1.18`；两个下拉框；页码 14 项、分隔线在 0 基 3/7/11，页眉 6 项、分隔线在 0 基 3，均无 `"-"`；截图 `v118_page_dropdown.png` 至少 3 条灰线、`v118_header_dropdown.png` 至少 1 条（颜色 `0xC8C8C8`，线长≈286 / 105 px）；键盘路径选中分隔条后跳到相邻正常项并写注册表，鼠标路径（下拉展开）退回原值、注册表不变；进程不崩；注册表测试前快照、测试后恢复。
- 机制验证（`W:\Eyes\Temp\opencode\test_sep_combo.ahk`）：`+0x210` 生效，实际样式 `0x50010213`；`DRAWITEMSTRUCT` 64 位偏移正确，12 个条目逐条收到 `WM_DRAWITEM`；3 条分隔线各 286 px。
- 剩余：用户 GUI 验收（实际点开两个下拉看分隔线、上下键能穿过）＋ 在新窗口把「在 PDF 文件右键菜单中添加…」取消再勾选一次（`RepairPdfContextMenu()` 只在旧 EXE 路径失效时切换）。

## 3. 已完成内容（v1.4，可视为稳定基线）

### 3.1 AHK 外壳 `PDF 加页码工具.ahk`（约 1869 行，UTF-8 带 BOM + LF）

- `PDF_TOOL_VERSION := "1.4"` → 决定 `%LOCALAPPDATA%\PdfPageNumberTool\backend-<版本>\` 目录名与 `backend.ok` 内容。
- `PDF_FORMAT_LIST` 9 项文案与 `6.8.19.ahk` 的 `PageFormatList` 一致，不含「无」，默认「加总页码 1 / n 居右」。
- `PdfFormatArgs(formatStr)`（约 987 行）用 `InStr` 把文案翻成 `--format plain|dash|total` + `--align left|center|right`；**改文案必须同步改它**。
- 注册表根键 `HKCU\Software\PdfPageNumberTool`（不要和 `DocFormatTool` 混用）：`PageNumFormat`（存文案）、`MarginRight`（唯一水平边距，名字为兼容保留）、`StartPage`、`FontSize`、`MarginBottom`、`CreateBAK`、`BakDir`、`ContextActive`。
- 原地写回顺序（不可改）：`IsPdfWritable()` → `BackupPdfFile()` → 后端写 `原名.pdf.pgnum.tmp` → `ReplaceOriginalPdf()`。
- 列表右键：打开文件 / 打开所在目录（`explorer /select`） / 恢复上一个版本（Yes-No 默认否 + `IsPdfWritable()`，不写日志） / 移除所选。底部「打开文件目录」按钮已删。
- 右键菜单文案「PDF 文件添加页码(&Z)」，键名 `PdfAddPageNum`，写 `.pdf\shell` 与 `SystemFileAssociations\.pdf\Shell` 两处；`RepairPdfContextMenu()` 会自愈旧文案与失效路径。

### 3.2 后端 `pdf_page_number.py`（14251 字节，360 行）

- `format_page_text(number, total, style)`、`make_overlay(..., align)`、`add_page_numbers(..., style, align)`；`--format` / `--align`。
- 定位：先算**视觉坐标** `xv`（left→margin、center→vw/2、right→vw-margin），再 `visual_to_content()` 逆旋转到内容坐标；旋转页 90/180/270 与不规则 CropBox 均正确。
- 实测坐标（A4、2cm/1.2cm）：left x=56.69、center x=297.64、right x=538.58，y=34.02；旋转 90° 页为 (561.26, 56.69/420.94/785.20)。

### 3.3 产物

- `build\backend.zip` 18.2 MB（v1.4 后端，PyInstaller onedir 打包）。
- `PDF 加页码工具.exe` 19.5 MB，2026-10-30 13:01 编译，图标 `pdf_page_icon.ico`。
- 文档：`AGENTS.md` 的 PDF 段、`程序整体说明.md` 第 13 节均已同步 v1.4。

## 4. 未完成 / 待修改内容

**v1.6 已完成（第 5.5 节），v1.7–v1.18 均已完成（v1.13–v1.18 待用户 GUI 验收），当前无阻塞待办。** 剩下的只是用户还没拍板的包体积/杀软方案（第 5.6 节）。

用户原话：「看到有几处还在弹出对话框，所有地方都不弹出对话框，全部使用右下角提示框，除非是要用户有所选择的，比如把文件夹拖进去，需要让用户选择是否检索子文件夹内的文件时。」
追加要求：「状态栏上：结果已同时在屏幕右下角提示。这样的文字就不要了，显得太傻了。」

排查结果（`grep MsgBox`，v1.5 共 9 处）：

| 行号 | 场景 | 处理 |
| --- | --- | --- |
| 412 | 拖入含子文件夹 → 询问是否递归 | **保留**（YesNo，必须用户选择） |
| 792 | 恢复上一个版本 → 覆盖确认 | **保留**（YesNo 默认否，破坏性操作必须确认） |
| 874 | `DirSelect` 选择备份文件夹 | **保留**（系统选择器，本身就是选择动作） |
| 488 | 列表为空 | 改浮动提示 |
| 499 / 504 / 509 / 514 | 起始页码 / 字号 / 水平边距 / 距下边 校验失败 | 改浮动提示 |
| 534 | `ResolveRunner()` 返回错误（后端无法启动） | 改浮动提示（错误文本多行，提示框高度自适应） |
| 553 | 列表里没有有效 PDF 路径 | 改浮动提示 |

其余反馈（正在处理、备份开关、右键菜单、恢复失败/成功、处理结果）在 v1.5 就已经是浮动提示，无需改动。

用户反馈原话：「界面再调整一下」+ 截图 `W:\Eyes\Temp\opencode\uploads\a4b4152c-18d7-471e-a7a1-4cea4ce9e769\image.png`（702×737，100% 缩放，客户端 700×705）。

### 已诊断出的真实缺陷（像素级实测，作为修改依据）

| 问题 | 证据 |
| --- | --- |
| 「水平边距(cm)」标签换行，`(cm)` 掉到下一行 | 声明 `x325 w75`，但该文字实测需 **77 px**；墨迹到 x=377 就断，`y=126–133` 处另有一块 `x=327–351` 的墨迹＝第二行 `(cm)`。AHK 的 Text 省略 `h` 时按内容自动撑高，于是压在两行控件之间 |
| 四个输入框间距不齐 | 标签末端→输入框的间距依次 14 / 20 / 34 / 21 px（34 就是换行那格） |
| 留白不均 | 「页码参数」组框 75–250，最后一行灰字到 203 → 底部空 44 px；「选项」组框 265–415 内容到 405 → 几乎顶边 |
| 右侧空旷 | 「页码格式」下拉 `x105 w300` 到 405 结束，405–680 一大片空白 |

## 5. 已完成的具体修改（v1.5，作为后续改动的坐标基线）

全部落在 `w:\eyes\Desktop\公文\PDF 加页码工具.ahk` 的 `ShowPdfGui()`（现为 258–374 行一带）。

### 5.1 版本号与修改记录（必须先做）

1. `PDF_TOOL_VERSION := "1.4"` → `"1.5"`（文件名**保持不变**，改名会让已注册右键路径与 EXE 失效）。
2. 文件头「修改记录」在 `;   v1.4 1. ...` 之前插入 v1.5 条目，内容要点：参数区改 4 列等距、标签右对齐防换行、页码格式下拉通栏、纵向重新分配、窗口 700×705→700×685；仅界面，无功能变更。
3. 文件头 `; 用法（v1.4 起）` 与正文说明无需大改，但可补一句 v1.5 只改界面。

### 5.2 参数区（`GroupBox`「页码参数」）

目标坐标（`w=700` 客户端，组框 `x20 w660`，可用内宽 35–665）：

| 控件 | 新坐标 |
| --- | --- |
| GroupBox | `x20 y75 w660 h140`（原 `h175`） |
| 标签「起始页码」 | `x35 y102 w85 Right` |
| Edit 起始页码 | `x125 y98 w55` |
| 标签「字号(pt)」 | `x195 y102 w85 Right` |
| Edit 字号 | `x285 y98 w55` |
| 标签「水平边距(cm)」 | `x355 y102 w85 Right` |
| Edit 水平边距 | `x445 y98 w55` |
| 标签「距下边(cm)」 | `x515 y102 w85 Right` |
| Edit 距下边 | `x605 y98 w55` |
| 标签「页码格式」 | `x35 y138 w85 Right` |
| DropDownList | `x125 y134 w540`（原 `x105 w300`） |
| 灰字说明 1（输出/原地写回） | `x35 y166 w630`（不变） |
| 灰字说明 2（透明图层/居中忽略边距） | `x35 y190 w630`（不变） |

要点：
- 标签统一 `w85` + `Right`（AHK 静态文本右对齐），四个 Edit 左边线才会齐；`水平边距(cm)` 77 px < 85 px，**不再换行**。
- 列起点 35 / 195 / 355 / 515，列间空 15 px；第 4 列 Edit 到 660，距组框右边 20 px。
- 组框底边 215，最后一行灰字到 203 → 底部留 12 px。

### 5.3 纵向重排

| 区域 | 新坐标 | 内部控件相对偏移（沿用原值） |
| --- | --- | --- |
| GroupBox「选项」 | `x20 y225 w660 h155` | 相对 +25 备份复选框、+53 备份目录行、+82 灰字、+102 右键复选框、+126 灰字 ⇒ 实际 y = 250 / 278 / 307 / 327 / 351 |
| GroupBox「PDF 任务列表」 | `x20 y390 w660 h215` | ListView `x30 y416 w640 h178`（原 `y451 h168`） |
| 底部按钮行 | `y615 h35` | 清空列表 `x20 w140`、移除所选 `x175 w140`、开始添加页码 `x340 w340` |
| 窗口 | `pdfGui.Show("w700 h685")`（原 `h705`） | 状态栏自动贴底，按钮底 650 与状态栏 661 之间 11 px |

### 5.4 验证与发布

1. `& 'C:\AutoHotkey\v2\AutoHotkey64.exe' /validate PDF 加页码工具.ahk` → 期望 `EXIT=0`。
2. 启动 GUI（无参数），截图窗口做墨迹分析，确认：参数行只有一条文字带（`y126–133` 不再有孤立墨迹）、四列 Edit 左边界等距、组框底部无大空白。截图/分析脚本见第 7 节。
3. 重编译 EXE（后端未变，**不需要重打 `build\backend.zip`**）：
   `Ahk2Exe.exe /in 'PDF 加页码工具.ahk' /out 'PDF 加页码工具.exe' /base 'C:\AutoHotkey\v2\AutoHotkey64.exe' /silent verbose /icon 'pdf_page_icon.ico'`
4. 同步文档：`AGENTS.md` 的 PDF 段（窗口尺寸、参数区布局、版本 1.5）、`程序整体说明.md` 第 13.5 节（窗口与右键菜单）与 13.6 节（编译顺序说明「仅改界面时不必重打 zip」）。
5. 启动 GUI 交给用户测试，回复末尾写「已完成」。
6. 回到本文件，把第 2/4/5 节更新为 v1.5 已完成，并把第 9 节补一条记录。

### 5.5 v1.6：把剩余模态框改成右下角提示（已完成 2026-10-30 14:07）

全部落在 `PDF 加页码工具.ahk` 的 `RunGuiProcessing()`（约 478–560 行），只替换 5 处 `MsgBox`，不动其它逻辑：

| 位置 | 改法 | 预期结果 |
| --- | --- | --- |
| 列表为空 | `uiPdfSB.SetText(" 列表里还没有 PDF，请先把 PDF 文件或文件夹拖进列表。")` + `ShowPdfToast("列表里还没有 PDF", "请先把 PDF 文件或文件夹拖进列表。", false, 3600)` | 不弹框；右下角红底提示 3.6 s，状态栏留一行 |
| 4 项参数校验 | 各自 `uiPdfSB.SetText(...)` + `ShowPdfToast("参数有误", "<具体原因>", false, 3600)`，仍保留 `Focus()` | 不弹框；提示 3.6 s，焦点落到出错的输入框 |
| `runner.exe = ""` | `uiPdfSB.SetText(" 无法启动后端：" 错误第一行)` + `ShowPdfToast("无法启动后端", runner.err, false, 6000)` | 长错误文本完整显示，提示框高度按内容自适应 |
| 无有效 PDF 路径 | `uiPdfSB.SetText(" 列表里没有有效的 PDF 路径。")` + `ShowPdfToast("没有可处理的 PDF", "列表里没有有效的 PDF 路径。", false, 3600)` | 不弹框 |

**状态栏文案规则（用户 14:10 追加要求）**：状态栏只写「发生了什么」，不写「结果已同时在屏幕右下角提示」这类元描述；
另外把 `OnPdfRunFinished()` 里 `处理完成：成功 X 个，失败 Y 个。结果已同时在屏幕右下角提示。` 的第二句删掉。

实施结果：`PDF_TOOL_VERSION` 已 1.6；`/validate` EXIT=0；用临时副本（`W:\Eyes\Temp\opencode\test_v16.ahk`）注入钩子依次触发“列表为空”“参数有误”，截取提示框与状态栏 OCR 确认：
提示框标题/正文正确（深红底 `5C2020`，360×72），状态栏显示「列表里还没有 PDF，请先把 PDF 文件或文件夹拖进列表。」「参数有误：起始页码必须是 1 到 999999 之间的整数。」，**没有元描述**。
EXE 已重编译（20,470,272 字节，14:07:03），静默跑副本 exit=0、耗时 2.76 s，自动解压 `backend-1.6` 并清理 `backend-1.5`，日志正常追加。

### 5.6 包体积与杀软误报（方案已定稿，**用户要求暂不动代码**）

> 2026-10-30 结论：用户选择「只出方案，暂不动代码」「暂不签名，先做精简」。
> 因此本节只存档实测数据与实施步骤，等用户明确说「开始做 A」时再照第 5.6.5 节执行。

#### 5.6.1 现状实测

> v1.14 基线：EXE 20,925,440 B（v1.13 为 21,303,296 B），内嵌 `backend.zip` 仍 19,551,127 B；下表 20,470,272 B 是 v1.5–v1.12 的旧基线，仅供对比。

| 项目 | 数值 |
| --- | --- |
| `PDF 加页码工具.exe` | 20,470,272 B（19.5 MB） |
| 其中 AHK 外壳部分（EXE − zip） | 1,366,175 B（1.3 MB） |
| 内嵌 `build\backend.zip` | 19,104,097 B（18.2 MB） |
| 后端解压后 | 35.07 MB / 33 个文件 |
| 后端 exe 本体（onedir 引导 + PYZ） | 5.92 MB |

体积大头：`PIL\_avif.cp314-win_amd64.pyd` **7.52 MB**、`python314.dll` 6.47 MB、`libcrypto-3.dll` **5.95 MB**、后端 exe 5.92 MB、`base_library.zip` 1.33 MB、`libssl-3.dll` 1.27 MB、`PIL\_imaging` 2.53 MB。

环境：PyInstaller 6.22.3 / Python 3.14.7 / reportlab 5.1.0 / pypdf 6.21.4。
签名：`PDF 加页码工具.exe` 与 `pdf_page_number.exe` 均 **NotSigned**；`C:\Python\python.exe`、`python314.dll` 由 Python Software Foundation 签名有效。

**关键判断**：`_avif`（7.52 MB）和 OpenSSL 全家桶（libcrypto + libssl ≈ 7.2 MB）本工具一行都不用——
后端只做「reportlab 画一行字 + pypdf 叠页」，不读图片、不做 TLS、不解密 PDF。它们是 PyInstaller 顺着依赖链自动抓进来的。

#### 5.6.2 三条路线（全部实测，不是估算）

| 方案 | 解压后 | backend.zip | 推得主 EXE | 状态 |
| --- | --- | --- | --- | --- |
| 现状 | 35.07 MB | 19,104,097 B | 19.5 MB | 基线 |
| 排除 PIL 整个包 | 23.54 MB | 12.06 MB | — | ❌ **运行失败** `No module named 'PIL'`（reportlab canvas 硬依赖 PIL） |
| **A. 排除无用模块** | 19.29 MB | **11,577,242 B（11.04 MB）** | **12,943,417 B（12.3 MB）** | ✅ 3 页 PDF 加页码 exit=0，页码坐标正常 |
| **B. A + 不用 reportlab/PIL** | **13.29 MB** | **7,305,111 B（6.97 MB）** | **8,671,286 B（8.3 MB）** | ⚠️ 体积已实测（build 时排除 reportlab/PIL），但功能未实现，需重写渲染 |
| C. B + 改分发形态 | — | — | EXE ≈1.5 MB + 后端目录 | 未实施 |

A 的具体排除项（已验证可用）：

```
--exclude-module PIL._avif --exclude-module PIL._webp --exclude-module PIL._imagingcms
--exclude-module PIL.ImageTk --exclude-module PIL.ImageQt
--exclude-module reportlab.graphics.renderPM --exclude-module reportlab.graphics.renderSVG --exclude-module reportlab.graphics.renderPS
--exclude-module ssl --exclude-module _ssl --exclude-module _hashlib --exclude-module cryptography
```

A 之后剩余大文件：`python314.dll` 6.47 MB、后端 exe 5.8 MB、`PIL\_imaging` 2.53 MB、`base_library.zip` 1.33 MB —— 已经接近 PyInstaller + reportlab 方案的下限。

#### 5.6.3 压缩方式实测

| 方式 | 体积 | 能否运行期解压 |
| --- | --- | --- |
| `backend.zip`（deflate 9） | 18.2 MB（现状） | ✅ 用系统 `tar.exe -xf` |
| `tar.xz`（preset 9） | 13.54 MB（压 A 前的原始后端） | ❌ Windows 自带 `tar.exe` 报 `Can't initialize filter`（没有 xz 解码器） |
| `tar.gz`（level 9） | 36.28 MB（反而更大） | ✅ 但没意义 |

结论：**运行期自解压只能继续用 zip**，压缩方式不是优化点。

#### 5.6.4 加密 PDF 兼容性验证（回答「排除 cryptography 安不安全」）

用 pypdf 的 RC4-128 加密样本分别喂给裁剪版后端（build C，已排除 `cryptography`/`ssl`/`_hashlib`）和源码版后端：

| 后端 | exit | status 文件内容 |
| --- | --- | --- |
| 裁剪版（无 crypto） | 2 | `PDF 已加密，需要密码，无法自动加页码` |
| 源码版（同环境） | 2 | `PDF 已加密，需要密码，无法自动加页码` |

**两者行为完全一致**：后端在读页之前就检出加密并友好退出，加密文件根本不进解析流程。
（AES-256 样本没造出来，因为本机未装 `cryptography`，`pypdf` 的 fallback provider 无法生成 AES 样本；不影响该结论。）

#### 5.6.5 方案 A 的实施步骤（等用户批准后照做）

1. 先按 AGENTS.md 的 PyInstaller 命令重打后端，附加 5.6.2 的 `--exclude-module`；
2. 再按原命令压 `build\backend.zip`；
3. 重编译 `PDF 加页码工具.exe`（`Ahk2Exe`，`Start-Process -Wait`）；
4. 验证：删掉 `%LOCALAPPDATA%\PdfPageNumberTool` 模拟首跑 → 静默跑 PDF 副本 → exit=0、坐标 56.69/297.64/538.58、y≈34.02、日志追加；再跑一次确认走快路径不再解压；
5. 同步 `AGENTS.md` 的编译命令段与第 5.6 节。

注意：**A 只改打包参数，`.ahk` 与 `pdf_page_number.py` 一行都不改**，所以不需要升版本号、不需要重打前端逻辑；但 `build\backend.zip` 与 EXE 会变，需重新验证首跑解压。

#### 5.6.6 杀软误报：真正的原因不是体积

体积只影响下载体验，拦截看行为特征。按性价比排序：

1. **代码签名**（OV/EV 证书）——最有效，同时消掉 SmartScreen「未知发布者」。用户当前选择「暂不签名」。
2. **去掉「运行期释放 exe 再执行」的 dropper 行为**——现状是：AHK EXE 里塞 18 MB 压缩包 → 首跑用 `tar.exe` 解到 `%LOCALAPPDATA%` → 执行刚落地的新 exe。这是启发式引擎最典型的画像，**比体积危险得多**。改便携文件夹（后端目录随包分发，不落地）或 Inno Setup 安装包（LZMA）可直接降一个量级。
3. **向微软提交误报**（免费）：<https://www.microsoft.com/en-us/wdsi/filesubmission>，每次发版提交一次。
4. **不要 UPX**（强特征，直接触发通用报毒）；给后端 exe 补版本资源/图标（`--version-file`）。
5. 长期可选：Python embeddable 替代 PyInstaller bootloader，或 Go/Rust/C# 重写后端（单文件 1–3 MB、无释放动作）。

**注意**：任何打包方案变更都会改 `build\backend.zip` 与编译流程；只改 `.ahk` 时才可跳过 PyInstaller。

#### 5.6.7 「把 Python 运行时一起打包」实测（2026-10-30）

思路：不用 PyInstaller，直接把官方 embeddable Python + 两个依赖库放进包里，运行时调 `python.exe pdf_page_number.py`。

| 阶段 | 解压后 | 文件数 | zip（deflate 9） |
| --- | --- | --- | --- |
| 官方 `python-3.14.7-embed-amd64.zip` 原始 | 23.51 MB | 37 | 12,673,227 B（12.09 MB） |
| + pypdf/reportlab/PIL(剪 `_avif`/`_webp`/`_imagingcms`)/charset_normalizer，删 TLS/sqlite/winsound 等 | **26.95 MB** | 424 | **14,941,615 B（14.25 MB）** |

结果：`python.exe` 直跑后端 **exit=0**，页码文字坐标与 PyInstaller 版完全一致（`(-20.43, 0)`）；单文件耗时 1.08–1.10 s（3 次）。
原始 embeddable 自带的体积：`python314.dll` 6.47 MB、`libcrypto-3.dll` 5.95 MB、`python314.zip`（标准库）3.95 MB、`sqlite3.dll` 1.51 MB、`libssl-3.dll` 1.27 MB。

**结论：这条路比方案 A 更大**（zip 14.25 MB vs 11.04 MB，推得主 EXE ≈15.6 MB vs 12.3 MB）。原因有两条：
1. `python314.zip` 是**完整标准库**（3.95 MB 已压缩），而 PyInstaller 只带真正 import 到的模块（`base_library.zip` 1.33 MB）；
2. `reportlab`/`pypdf` 的 `.py` 源码和 `reportlab\fonts`（`.afm`/`.sfd`，含 0.5 MB 的 DarkGarden）全在，PyInstaller 会把它们编成字节码并只留必要资源。

**但它有一个真正的价值**：`python.exe`/`python314.dll` 是 Python Software Foundation **签名**的，且不需要 PyInstaller bootloader。
所以它适合作为「便携文件夹」里的运行时（第 5.6.6 节第 2 条的方案 C），不适合再塞回单文件 EXE 去自解压。

#### 5.6.8 aardio 评估（2026-10-30，本机 `W:\aardio` 已装 v45）

用户提问：「可以用 aardio 写吗，它可以打包 Python 环境，不需要释放 EXE。」

核查结论（查 aardio 官方文档 + 社区教程 + 本机安装目录）：

1. **aardio 自身确实免释放**：它的文件函数直接读 EXE 内嵌资源（`/res/...`），纯 aardio 程序可以真正做到单文件不落地。
2. **但 Python 部分做不到**：aardio + Python 必须携带解释器，`.py` 一般也放在硬盘上（社区教程原话）。
   所以只有两条路：① 便携文件夹（不再单文件）；② 首跑把 Python 目录解压到 `%LOCALAPPDATA%`/`%TEMP%`（**dropper 行为照旧**，只是解压代码从 PyInstaller 换成自己的）。
3. **体积不会更小**：aardio 外壳约 0.6–1 MB，比 PyInstaller 引导（1.4 MB）只省不到 1 MB；Python 运行时仍要按第 5.6.7 节那份算（≈27 MB 解压 / 14 MB 压缩），最终 ≥ 方案 A，比方案 B 大得多。
4. **杀软不保证**：aardio 官方《如何解决软件发布或误报问题》自己写明「Defender 基本不认识的 EXE 就会误报」「不要以为改几句代码，或者换个编程语言没有误报就安全了」；aardio 被大量外挂/灰产使用，360/火绒都有对应启发式。唯一的好处是落地的是签名过的 `python314.dll`，不是来历不明的 bootloader。
5. **成本最大**：要把 1714 行 AHK 前端（GUI、ListView 拖放、右键菜单、托盘、注册表、toast、日志、备份/恢复）整体重写成 aardio，还要换构建链（`W:\aardio\aardio.exe` v45 是 IDE，命令行发布支持弱，现在 `Ahk2Exe` 那套可脚本化编译会丢）。等于把「改一个 `.py`」变成「重写整个工具 + 全量回归」。

**排序结论**：要小 → 方案 B（zip 6.97 MB / EXE 8.3 MB）；要去掉 dropper 特征 → 方案 C（便携文件夹或 Inno 安装包，运行时用第 5.6.7 节的 embeddable 目录）；要单文件 + 小 + 零释放 → 只能走原生重写（Go/Rust/C#），不是 aardio。
aardio 只有在「决定整体换技术栈」时才值得考虑，作为体积/杀软优化手段性价比最低。

## 6. 关键文件与目录说明

| 路径 | 说明 |
| --- | --- |
| `PDF 加页码工具 v1.18.ahk` | 当前外壳源码（2596 行，UTF-8 带 BOM + LF）。**文件名带版本号**，升版本新建文件；`v1.13`–`v1.17` 保留作回退 |
| `PDF 加页码工具.ahk` / `PDF 加页码工具.exe` | v1.12 成品，保留作回退，不再修改 |
| `pdf_page_number.py` | 后端源码（1452 行）；`--format/--align` + `--header*` + 三态开关 `--no-page-numbers`/`--keep-page-numbers`/`--no-header`/`--keep-header`；改动后必须重打 zip 再编译 |
| `pdf_page_icon.ico` | EXE 图标（编译参数 `/icon`） |
| `build\spec\pdf_page_number.spec` | PyInstaller 配置（onedir + noconsole） |
| `build\dist\pdf_page_number\` | PyInstaller 产物（`pdf_page_number.exe` + `_internal`） |
| `build\backend.zip` | 由上述目录压缩，`FileInstall("build\backend.zip", …)` 编译期读取（v1.17 为 19,123,458 B / 33 项） |
| `PDF 加页码工具 v1.18.exe` | 编译产物（20,505,088 B），已被 `.gitignore` 排除；`v1.13`–`v1.17.exe` 保留作回退 |
| `PDF加页码执行日志.txt` | 运行日志，追加不轮转：`时间<Tab>结果<Tab>路径<Tab>说明` |
| `%LOCALAPPDATA%\PdfPageNumberTool\backend-1.18\` | 首跑解压目录，带 `backend.ok`（内容＝版本号） |
| `我的文档\PDF加页码文件备份\` | 默认备份目录（本机 `w:\eyes\Documents\PDF加页码文件备份`） |
| `W:\Eyes\Temp\opencode\` | 临时测试区（PDF 副本、测试脚本、OCR 结果都放这里） |

## 7. 运行、测试与验证命令

```powershell
# 界面（无参数 → 建窗口）
Start-Process 'C:\AutoHotkey\v2\AutoHotkey64.exe' -ArgumentList '"W:\eyes\Desktop\公文\PDF 加页码工具 v1.18.ahk"'

# 语法校验（EXIT=0 为通过）
cmd /c "C:\AutoHotkey\v2\AutoHotkey64.exe /validate ""PDF 加页码工具 v1.18.ahk"" & echo EXIT=%errorlevel%"

# 源码态静默处理（会真的改文件、写备份、写日志，务必用副本）
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = 'C:\AutoHotkey\v2\AutoHotkey64.exe'
$psi.Arguments = "/ErrorStdOut `"W:\eyes\Desktop\公文\PDF 加页码工具 v1.18.ahk`" --quiet `"<PDF副本路径>`""
$psi.UseShellExecute = $false; $psi.RedirectStandardError = $true
$proc = [System.Diagnostics.Process]::Start($psi); $proc.WaitForExit(120000) | Out-Null
$proc.ExitCode

# 重打后端（仅在改了 .py 时需要；源码态调试用不到）
& 'C:\Python\Scripts\pyinstaller.exe' --noconfirm --clean --onedir --noconsole --name pdf_page_number --distpath build\dist --workpath build\work --specpath build\spec 'pdf_page_number.py'
& 'C:\Python\python.exe' -c "import os,zipfile; src=r'build\dist\pdf_page_number'; z=zipfile.ZipFile(r'build\backend.zip','w',zipfile.ZIP_DEFLATED,compresslevel=9); [z.write(os.path.join(dp,f), os.path.relpath(os.path.join(dp,f),src).replace(os.sep,'/')) for dp,dn,fn in os.walk(src) for f in fn]; z.close()"

# 编译 EXE（GUI 子系统，必须 -Wait；v1.13 起文件名带版本号）
Start-Process -FilePath 'C:\AutoHotkey\Compiler\Ahk2Exe.exe' -ArgumentList @('/in',(Resolve-Path 'PDF 加页码工具 v1.18.ahk').Path,'/out',(Join-Path (Get-Location) 'PDF 加页码工具 v1.18.exe'),'/base','C:\AutoHotkey\v2\AutoHotkey64.exe','/silent','verbose','/icon',(Resolve-Path 'pdf_page_icon.ico').Path) -Wait
```

坐标核对：用 `pypdf` 的 `extract_text(visitor_text=…)` 取 `Tm`，**再乘 `CTM`** 才是真实坐标（只读 `tm` 会得到 0）。A4、2cm/1.2cm、10.5pt、居左时，页码基线 y≈34.02、页眉基线 y≈807.9，两者 x≈56.69；居中/居右的 x≈297.64 / 538.58。

界面像素校验（模型读不了图，用这个方法代替「看」）：

```powershell
# 1) 截窗口位图 → 2) 逐行统计墨迹 → 3) 看参数行是否还有孤立墨迹 / 各列间距
& 'C:\Python\python.exe' 'W:\Eyes\Temp\opencode\ink_rows.py'   # 需要先把截图路径改成新文件
& 'C:\Python\python.exe' 'W:\Eyes\Temp\opencode\ink_cols.py'
# OCR（Windows 自带，zh-Hans 可用）：参考 W:\Eyes\Temp\opencode\ocr3.txt 的生成方式
```

## 8. 已知问题、风险与注意事项

1. **模型不能读图**：用户贴的截图必须靠 OCR（`Windows.Media.Ocr`，zh-Hans-CN 可用）＋像素墨迹分析来「看」，不要假装看到内容。3 倍放大后再 OCR（`PIL` 的 bicubic）识别率高得多。
2. **`edit`/`write`/`read` 工具写不进盘**（之前实测：写入未落盘）。所有改文件用 PowerShell 或 Python 脚本 `Path.write_text(...)`，按目标文件原行尾写（本文件 CRLF、`.ahk`/`.py`/`程序整体说明.md` 是 LF），改完用 Python 读回校验；中文控制台输出会乱码，**写进文件再用 read 工具看**。
3. **编码与换行**：当前 `PDF 加页码工具 v1.18.ahk` 是 UTF-8 **带 BOM** + LF（2596 行）；`pdf_page_number.py` 是 UTF-8 无 BOM + LF（1452 行）。改完必须校验 BOM/行尾，别引入 CRLF 或丢 BOM。
4. **版本号会触发后端重解压**：每次改 `PDF_TOOL_VERSION`（如 1.12→1.13），首次运行会在 `%LOCALAPPDATA%` 重新解压（约 2.5 s）并尝试清理其它 `backend-*`。只改 `.ahk` 且后端 zip 没变时，不需要重打 zip，只重编译 EXE。
5. **编译顺序**：`FileInstall` 的 `build\backend.zip` 是编译期读取；只改 `.ahk` 时可以跳过 PyInstaller。`Ahk2Exe.exe` 是 GUI 子系统，PowerShell `&` 调用不会等待，要用 `Start-Process -Wait` 或 `.NET Process.WaitForExit()`。
6. **`Start-Process -Wait -PassThru` 的 `ExitCode` 可能为空**，用 `[System.Diagnostics.Process]::Start()` + `WaitForExit()` 更可靠。
7. **文案与解析耦合**：`PDF_FORMAT_LIST` 文案 ←→ `PdfFormatArgs()`（`InStr`）←→ 后端 `--format/--align`，改任一处必须同步其余两处；注册表 `PageNumFormat` 存文案，不认识就回落默认值。v1.17 的两个特殊项「不改变页码」「清空页码」由 `ProcessOne()` 先判哨兵文案，不走 `PdfFormatArgs()` 的结果；页眉同理（`PDF_HEADER_KEEP` / `PDF_HEADER_CLEAR`）。
8. **注册表键名不能改名**：`MarginRight` 与 `--margin-right` 的名字为兼容旧版保留，含义已是「唯一水平边距」。
9. **右键菜单两处都要写**：`.pdf\shell` 与 `SystemFileAssociations\.pdf\Shell`，否则部分系统只显示一处。
10. **`Static` 假按钮**：「开始添加页码」是 `Text` 控件（`+0x200 BackgroundE8F5E9 Border`），`BM_CLICK` 无效，自动化测试要用鼠标消息或脚本内钩子。
11. **AHK Text 控件不给 `h` 会按内容撑高**：标签宽度只要差 1–2 px 就会换行并被误认为「被遮挡」，定宽度前先量文字（10 号微软雅黑一个汉字约 13.3 px）。
12. **仓库是 git 仓库**（远程 `ksgujie/gongwen`，分支 `main`），但工作区里 20 个旧版本显示为已删除、当前脚本未跟踪；**不要把 `git add -A` 当回退手段**，改前仍先把 `.ahk` 复制一份到 `W:\Eyes\Temp\opencode\`。
13. **页眉会与原文重叠**：页眉直接叠加在最上层，原文顶部 1.2 cm 内有内容就会压在一起；这是已知限制，不做自动避让。
14. **升版本后右键菜单不会自动切换**：`RepairPdfContextMenu()` 只在旧路径失效时重写；旧 EXE 还在时，必须在新窗口里取消再勾选一次。
15. **页眉空串读取**：`HeaderText` 允许空串，必须用 `GetPdfRawStr()`；`GetPdfStr()` 会回落默认值。
16. **旧层清理禁止按字号扫全流**：v1.10–v1.12 的兜底会误删 10.5 pt 正文；v1.13 以 `DocFmtPN` 标记优先。
17. **不要改回“后端直接写原文件旁边”（v1.14）**：沙箱/网盘重定向目录里冻结后端的写入会落到另一层视图，外壳看不见（本机 Sandboxie `OpenFilePath=w:\Eyes\Desktop\公文` 必现）。写回必须走 `%TEMP%` 中转，见第 2.3 节与 `ProcessOne()` ③④ 的注释。
18. **下拉框分隔线是自绘的（v1.18）**：两个下拉必须保留 `+0x210`（`CBS_OWNERDRAWFIXED|CBS_HASSTRINGS`），否则系统按文本项自己画、`WM_DRAWITEM` 不会来；`PdfSepDraw()` 里的 `DRAWITEMSTRUCT` 偏移是 **64 位**的（32 位不同：`hwndItem` 20、`hDC` 24、`RECT` 28、`itemData` 44）。分隔条只存在于 `PDF_FORMAT_ITEMS` / `PDF_HEADER_ITEMS`，**不要**塞进 `PDF_FORMAT_LIST` / `PDF_HEADER_ALIGN_LIST`（那两个还用于注册表白名单校验和 `ProcessOne()` 传参）。想改分组位置只动 `PdfAddSeps()` 的第二个参数（页码 `[3, 6, 9]`、页眉 `[3]`）。

## 8.1 AI 对 `PDF 加页码工具.ahk` 的代码审查发现（2026-10-30，已归档；v1.7 已据此实施，2026-11-02 已落地）

审查范围：`PDF 加页码工具.ahk` 全文（1714 行）+ `pdf_page_number.py`（360 行）；核对来源：AHK v2 官方文档（FileOpen flags / #SingleInstance / ListView.GetNext）与本机 pypdf 源码（`clone_document_from_reader` 不保留加密字典）。

结论：**备份→临时文件→替换、`backend.ok`、状态文件 UTF-8、日志先写失败的顺序设计是对的，不建议改动。**

2026-11-02：用户批准后，P0-1~P0-5 已全部落地到 `PDF 加页码工具.ahk`（单实例 Off、`FileExist` 检查、FileMove 重试、恢复快照、`NormalizePdfParams`），工具版本升为 v1.7 并已编译通过。

| 编号 | 严重度 | 位置 | 发现 | 建议 |
| --- | --- | --- | --- | --- |
| P0-1 | 高 | `PDF 加页码工具.ahk` 第 2 行 `#SingleInstance Force` | 与 `%1` 右键逐文件启动、`GUI 打开 + 右键` 等场景冲突；可能导致多选只处理一个，或 GUI 被新实例替换掉 | 改为 Off + 任务队列/锁 |
| P0-2 | 高 | `PDF 加页码工具.ahk` 第 1312 行 `IsPdfWritable` | `FileOpen(path,'rw')` 在目标不存在时会**新建**，缺失文件会过检查并留 0 字节文件/备份，且错误文案误导 | 先 `FileExist`，分开报错 |
| P0-3 | 高 | `PDF 加页码工具.ahk` 第 1373 行 `ReplaceOriginalPdf` | 兜底 `FileCopy` 非原子，可能半文件；覆盖顺序仍可接受但兜底安全不够 | FileMove 重试 + 替换后校验 |
| P0-4 | 高 | `PDF 加页码工具.ahk` 第 820 行 恢复覆盖 | 覆盖无当前备份、非原子 | 先备份当前再临时文件+替换 |
| P0-5 | 高 | `PDF 加页码工具.ahk` 第 1038 / 1205 行 | 静默/右键模式不校验注册表参数；`IsNumber` 放过小数 | 抽 `NormalizePdfParams()`，加载+处理前统一取整/限范围 |

**P1/P2 候选（未纳入 v1.7，用户未定）**：重复加页码叠加标记；加密/签名前后许可丢失；RunWait 无超时/取消；状态文件撞名/`EnsureBackend` 竞争；Toast「Iconi」被当成功；无添加文件按钮；无打开日志入口；Static 假按钮无键盘焦点；Toast 固定宽/主屏定位；备份目录每键写注册表；格式存文案；`#Warn All, Off`；无自动化测试。

## 9. 最近变更记录

| 时间 | 版本 | 内容 |
| --- | --- | --- |
| 2026-11-03 13:54 | v1.18 | 下拉框分组：页码下拉在「普通数字 / 格式数字 / 加总页码 / 不改变·清空」之间、页眉下拉在「对齐三项 / 不改变·清空」之间加分隔线；Win32 下拉框没有原生分隔线，改 owner-draw 自绘（`+0x210` + `WM_DRAWITEM`），占位项 `PDF_SEP` 只画 1px 灰线，`PdfSepGuard()` 把选中项跳到相邻正常项（鼠标点击退回原值），不会写进注册表；删掉页眉列表遗留的 `"-"`；只改 AHK 外壳（py 与 zip 不动），重编译 EXE（20,505,088 B）；`/validate` EXIT=0、EXE 静默 11 页副本 exit=0、GUI 16 项检查 0 失败
| 2026-11-03 13:01 | v1.17 | 在 v1.16 上做三态：页码下拉加「不改变页码」「清空页码」，页眉下拉改「不改变页眉/居左/居中/居右/清空页眉」，默认两项「不改变」；`ApplyPdfUiState()` 统一置灰，`RunGuiProcessing()` 只校验本次用到的参数，两项都「不改变」GUI 提示返回、静默/右键在备份前记失败；后端新增 `--keep-page-numbers`/`--keep-header`、独立指纹块 1.33/1.32（旧合并层 1.31 按视觉 y 拆分）、文字层原生页码/页眉识别（页脚/页顶 4cm 带）；重打 `build\backend.zip`（19,123,458 B / 33 项）、重编译 EXE（20,501,504 B）；`/validate` EXIT=0、后端 78 项检查 0 失败、源码/EXE 静默 7 组组合通过；扫描图片里的页码/页眉删不掉 |
| 2026-11-03 | v1.16 | 以 v1.14 为基线做界面调整：备份目录行输入框/选择/打开统一 `h28` 同一 `y`，打开改 28×28 系统图标按钮 + 悬停提示；页眉下拉首项「不添加页眉」，选它则页眉文字/距上边置灰并传 `--no-header`，标签改「页眉」、组框改「页眉（可选）」，默认值「不添加页眉」；只改 AHK 外壳（后端复用已含 `--no-header` 的 zip），EXE 20,927,488 B；`/validate`、源码/EXE 静默两条路径均 exit=0 |

| 2026-11-03 | v1.14 | 修复编译版在 `w:\eyes\Desktop\公文` 里必定失败（Sandboxie `OpenFilePath` 目录下冻结后端写入落到另一层视图，外壳 `FileExist()` 看不到）；`ProcessOne()` 改为“后端写 `%TEMP%` → 外壳复制回原目录 `.pgnum.tmp` → 校验大小 → 原子改名”，失败信息补退出码/状态/临时文件路径；只改 AHK 外壳，重编译 EXE（20,925,440 B）；`/validate` EXIT=0，源码态/EXE 静默/EXE 窗口三条路径均成功 |
| 2026-11-02 | v1.13 | 新增页眉（空=文件名、三对齐、距上边 0–20、与页码共用字号）；后端一次 overlay 画页眉+页码并注册中文字体；旧层清理改为 `DocFmtPN` 标记优先；程序文件名带版本号；后端重打 zip（19,551,127 B）、EXE 重编译（21,303,296 B）；`/validate` 与源码/EXE 静默跑副本均 exit=0 |
| 2026-11-02 | v1.7 | 用户批准「可以编写代码了」后实施 P0-1~P0-5：单实例 `Force`→`Off`；`ProcessOne` 缺文件检查；`IsPdfWritable` 前置 `FileExist`；`ReplaceOriginalPdf` FileMove 重试；恢复前快照；新增 `NormalizePdfParams`；`/validate` EXIT=0；重编译 EXE；`AGENTS.md` / `程序整体说明.md` / 本文件同步 |
| 2026-10-30 16:50 | v1.7（计划/待批准） | AI 代码审查发现归档第 8.1 节；P0-1~P0-5 形成 v1.7 实施计划，等待用户明确批准「可以编写代码」才启动 |
| 2026-10-30 13:01 | v1.4 | 重打后端 zip、用 `pdf_page_icon.ico` 重编译 EXE（19.5 MB）；右键菜单两处改名为「PDF 文件添加页码(&Z)」；`AGENTS.md`、`程序整体说明.md` 同步；静默跑「格式数字 — 1 — 居左」验证文字落在 x=56.69 |
| 2026-10-30 13:20 | v1.4 | 像素分析用户截图，定位「水平边距(cm)」标签换行等 4 处界面缺陷；用户批准 v1.5 界面重排方案 |
| 2026-10-30 13:30 | — | 建立本交接文件 |
| 2026-10-30 13:44 | v1.5 | 按方案改 `ShowPdfGui()` 坐标（4 列等距、标签 `w85 Right`、下拉 `w540`、组框 10px 间隔、窗口 700×685），`/validate` EXIT=0；截取客户端 700×685 位图＋墨迹分析确认：`y126–133` 的孤立墨迹消失，四个 Edit 左边界 125/285/445/605 等距，`水平边距(cm)` 单行落在 362–438 |
| 2026-10-30 14:07 | v1.6 | 5 处 `MsgBox` 改 `ShowPdfToast`（列表空、4 项参数、后端启动失败、无有效路径），状态栏只写事实、删掉“结果已同时在右下角提示”；`/validate` EXIT=0；临时副本钩子截图 + OCR 验证两条错误路径；重编译 EXE（20,470,272 B，14:07），静默跑 exit=0、2.76 s、`backend-1.6`；`AGENTS.md` / `程序整体说明.md` 同步 |
| 2026-10-30 14:20 | — | 实测打包方案：排除 PIL 会运行失败；排除 `_avif`/`ssl`/`crypto` 后 zip 18.2→11.04 MB 且可用；`tar.xz` 13.54 MB 但 Windows tar 不解 xz。数据写入第 5.6 节，等用户选方向 |
| 2026-10-30 15:10 | — | 用户选择「只出方案、暂不动代码」「暂不签名、先做精简」；补测方案 B 体积（解压 13.29 MB / zip 6.97 MB / 主 EXE ≈8.3 MB）与加密 PDF 兼容性（裁剪版与源码版同为 exit=2、同一提示，排除 crypto 行为中性）；第 5.6 节改写为定稿方案 |
| 2026-10-30 15:40 | — | 实测「打包 Python 运行时」：官方 embeddable + 依赖 = 26.95 MB 解压 / zip 14.25 MB（比 A 大），python.exe 直跑 exit=0、坐标一致；评估 aardio v45：自身免释放但 Python 仍须落盘，体积不占优，官方文档自认换语言不能免误报，且需重写 1714 行前端。写入 5.6.7 / 5.6.8 |
| 2026-10-30 14:05 | v1.6 | 用户要求「除必须选择的地方外不要弹对话框」。排查出 9 处 `MsgBox`：3 处保留（递归询问、覆盖确认、目录选择），6 处改浮动提示（列表空、4 项参数校验、后端启动失败、无有效路径）；计划写入第 5.5 节 |
| 2026-10-30 13:50 | v1.5 | 重编译 `PDF 加页码工具.exe`（19.5 MB，图标不变）；EXE 静默跑副本 exit=0，自动解压 `backend-1.5`、清理 `backend-1.4`，日志追加一行；`AGENTS.md`（版本 1.5 + 界面布局 + 标签宽度陷阱）、`程序整体说明.md`（13 标题、13.5 布局、13.6 只改 AHK 可跳过打包、13.7 坐标陷阱）已同步 |

## 10. 给下一个 AI 账号的接续指令

1. 先读第 2、4、8、9 节，复述：「当前是 v1.18（在 v1.17 三态基础上加下拉框分组分隔线）：页码/页眉各「添加/不改变/清空」，清空会删本工具旧叠加层和文字层里可识别的原生页码/页眉（扫描图片里的删不掉）；两个下拉改成 owner-draw 自绘，`PDF_SEP` 占位项只画 1px 灰线，`PdfSepGuard()` 保证它不会被选中；只改 AHK 外壳、没重打 zip，重编译 EXE（20,505,088 B），`/validate`、EXE 静默 11 页副本、GUI 16 项检查均通过，等用户做 GUI 验收并重勾右键菜单；v1.15/v1.16 是并行回退分支，不要合并；之后只剩第 5.6 节的包体积/杀软方案待拍板」。
2. 如果用户提出新的界面调整：先按第 5 节的坐标基线改 `ShowPdfGui()`，再用第 7 节的「界面像素校验」流程验证——**不要只凭代码推断改好了**（模型读不了图，必须靠截图墨迹分析/OCR）。
3. 如果用户报告功能问题：后端在 `pdf_page_number.py`，前端在 `PDF 加页码工具 v1.18.ahk` 的 `ProcessOne()` / `RunGuiProcessing()`；下拉框自绘在 `PdfSepDraw()` / `PdfSepGuard()`；改后端必须重打 `build\backend.zip` 再编译。
4. 任何改动前先更新本文件（第 2/4/9 节），再动手；改完更新第 2 节状态。**升版本时新建带版本号的 `.ahk`/`.exe`，并提醒用户重新勾选右键菜单。**
5. 回复用中文，确实改了代码时末尾写「已完成」。
6. 新功能/新版本必须先向用户要「可以编写代码了」的明确许可（协作约定：先规划，用户同意后才写代码）。
7. 下一步优先级：① 用户验收 v1.18（点开两个下拉，确认组间分隔线显示正常；键盘上下键能穿过分隔线、不会停在线上；选中后注册表只存正常文案；批量处理时下拉置灰仍能看清文字；右键菜单需在新窗口取消再勾选一次）；② 验收通过后再讨论第 5.6 节包体积/杀软方案。
