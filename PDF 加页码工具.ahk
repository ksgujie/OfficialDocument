#Requires AutoHotkey v2.0
; v1.7：#SingleInstance Force 改为 Off。
; 背景：Force 会“替换”旧实例。对本工具有害——
;   · 资源管理器右键多选 PDF 时 Explorer 对 "%1" 逐文件启动，第二个实例会杀第一个，
;     可能导致多选只处理到其中一个（或前几个刚完成就被杀）；
;   · 窗口打开着时右键一个 PDF，新静默实例启动会把 GUI 实例当旧实例替换掉，正在批次会被中断。
; Off 后多实例共存；每个实例只管一个文件（静默）或一个窗口（GUI）。
#SingleInstance Off
; 关闭警告弹窗：默认设置下，“变量可能未赋值”之类的警告会弹窗并卡住拖放处理，
; 对最终用户没有意义（语法错误仍会正常报错）。
#Warn All, Off
;=====================================================================
; PDF 加页码（AutoHotkey v2 外壳 + 图形界面）
;
; 作用：
;   把 PDF 文件（或装着 PDF 的文件夹）拖进窗口，直接在原文件上叠加页码
;   （v1.3 起原地写回），每页默认右下角显示“当前页 / 总页数”。
;   原 PDF 内容一个字节都不改（页码是叠加上去的新图层），所以不会重排版面，
;   扫描件也能加。v1.4 起页码格式与公文排版一致，可选：
;       普通数字 1 / 格式数字 — 1 — / 加总页码 1 / n，各支持居左、居中、居右。
;
;   真正的加页码工作在 pdf_page_number.py 里（pypdf + reportlab），
;   AHK 只负责：收集拖放的文件、拼命令行、调用后端、把结果汇总显示。
;
; 依赖（v1.2 起分两种运行方式）：
;   ① 编译版（PDF 加页码工具.exe）：后端已经用 FileInstall 嵌进 EXE，首次运行时自动
;      解压到 %LOCALAPPDATA%\PdfPageNumberTool\backend-<版本>\，之后直接复用。
;      用户机器不需要安装 Python，也不需要 pypdf / reportlab。
;   ② 源码态（直接用 AutoHotkey 跑本 .ahk）：仍走本机 Python + 同目录
;      pdf_page_number.py，便于改代码时免打包调试。
;
; 用法（v1.4 起）：
;   ① 双击（无参数）：打开窗口，把 PDF 或文件夹拖进列表，选好页码格式和参数后
;      点“开始添加页码”
;   ② 拖放：把 PDF/文件夹拖到本脚本（或编译后的 EXE）图标上 → 不显示窗口，
;      直接处理，处理完只在屏幕右下角弹一个提示框
;   ③ 右键：窗口里勾选“在 PDF 文件右键菜单中添加…”后，可在资源管理器右键 PDF
;      直接加页码（同样不显示窗口）
;   ④ 命令行：AutoHotkey64.exe "PDF 加页码工具.ahk" "a.pdf" "某文件夹"
;      加 --quiet 则完全不提示，只用退出码表示成败（0 全成功 / 1 有失败），
;      适合批处理、计划任务。
;   v1.3 起页码直接写回原文件（不再生成“_加页码.pdf”），写之前先备份到
;   “我的文档\PDF加页码文件备份\原名_yyyyMMdd_HHmmss.pdf”，备份失败就不动原文件。
;   文件夹在窗口模式会询问是否递归；拖到 EXE 的静默模式只扫描第一层。
;   ~$ 开头的 Office 锁文件会自动跳过。
;
; 修改记录：
;   v1.12 1. 修复批量处理必然崩溃：v1.8 把新按钮 uiPdfBakOpen 放进 SetPdfControlsLocked() 的
;           锁定列表，却漏了 global 声明，一点“开始添加页码”就报“局部变量未赋值”。
;        2. 静默模式补上重复文件计数：AddPdfFile() 遇到重复文件时 skipped += 1，
;           汇总提示才会出现“另有 N 个文件被跳过”（此前恒为 0）。
;        3. “没有找到可处理的 PDF 文件”改回失败样式：Notify 的图标参数 "Iconi" → "Icon!"。
;        4. 源码态注册右键菜单改用现文件名（原写旧名 PDF加页码.exe，改名后成品 EXE 被忽略，
;           退回 AutoHotkey + .ahk 路径，依赖本机 Python）。
;        仅改 AHK 外壳，后端与 build\backend.zip 未变，重编译 EXE 即可。
;   v1.11 1. 修复文件头修改记录乱码并修正备份目录行“打开”按钮文案（仅注释与界面文案，不改业务逻辑）。
;   v1.10 1. 清除已有旧页码：当源 PDF 已有页码时，后端先剥离尾部旧的页码叠加层，再叠加新页码，避免旧页码残留。
;        后端 `pdf_page_number.py` 已变更，必须重新打包 `build\backend.zip` 并重编译 EXE。
;   v1.9 1. 页码参数修改后立即保存到注册表（控件 `Change` → `SavePdfParamNow`），下次启动由 `LoadPdfSettings` 回填。
;   v1.8 1. 备份目录行新增“打开”按钮（`OpenPdfBakFolder`），可直接打开当前备份目录。
;   v1.7 1. 多实例策略修复：`#SingleInstance Force` 改为 `Off`，修复双选多个 PDF 时资源管理器逐文件启动导致实例被替换而只处理部分文件的问题。
;       2. `ProcessOne` 先单独判断 `FileExist`，再走 `IsPdfWritable`，避免文件不存在时 `FileOpen("rw")` “新建”空文件并误生成空备份。
;       3. `ReplaceOriginalPdf` 增加 `FileMove` 重试（多次、间隔 60 ms），先尝试原子改名，全部失败再退回“复制覆盖 + 删除临时文件”。
;       4. 恢复上一版本前先备份当前文件为 `_before-restore_<时间戳>`，快照失败则中止恢复，绝不直接覆盖。
;       5. 新增 `NormalizePdfParams`，加载注册表时规范化起始页码、字号、水平边距、距下边距，非 GUI 路径也保证参数不越界。
;       后端 `pdf_page_number.py` 与打包产物不变，只需重编译 EXE；界面、命令协议未动。
;   v1.6 1. 提示方式统一：除了“必须由用户做选择”的场景，一律不再弹模态对话框，改走
;           右下角浮动提示（ShowPdfToast），并把原因写进状态栏（提示框消失后仍能看到）。
;           本次改掉的 5 处：列表为空、4 项参数校验、后端无法启动、列表里没有有效路径；
;           同时删掉状态栏里的“结果已同时在屏幕右下角提示。”这句元描述。
;           仍保留对话框的 3 处（都是必须由用户做选择才能继续）：
;             · 拖入含子文件夹的文件夹时询问是否递归扫描（YesNo，默认“是”）；
;             · 右键“恢复上一个版本”的覆盖确认（YesNo，默认“否”，破坏性操作）；
;             · “选择…”按钮弹出的系统目录选择器（DirSelect，本身就是一个选择动作）。
;   v1.5 1. 界面重排（**只改外观，功能与命令行一律未动**）：
;          ① 参数区四个输入框改成等距四列（列起点 35/195/355/515，输入框统一 w55），
;             标签统一 w85 且右对齐——原来“水平边距(cm)”只给了 75px，而这句话在 10 号
;             微软雅黑下要 77px，AHK 的 Text 不给 h 时会按内容自动撑高，于是 (cm) 被挤到
;             第二行、压在参数行和页码格式行之间，看起来像被遮挡；
;          ② “页码格式”下拉改为通栏 w540（原来 w300，右侧留了 275px 空白）；
;          ③ 各组框纵向重新分配：页码参数 h175→140、选项 y265→225、任务列表 y425→390、
;             按钮行 y640→615，组框之间统一 10px 间隔，去掉页码参数组框底部 44px 空白；
;          ④ 窗口 700×705 → 700×685。
;   v1.4 1. 新增“页码格式”选项，可选 9 种（普通数字 / 格式数字 — 1 — / 加总页码 1 / n
;           × 居左 / 居中 / 居右），文案与公文排版一致，默认“加总页码 1 / n 居右”；
;           “距右边”改名“水平边距”：居右时距右、居左时距左、居中固定在页心；
;        2. 后端 pdf_page_number.py 新增 --format / --align，并按对齐方式改用
;           drawString / drawCentredString / drawRightString（旋转页仍走 visual_to_content）；
;        3. 右键菜单文字改为“PDF 文件添加页码(&Z)”，已注册旧文案的会在编译版启动时自愈；
;        4. 去掉底部“打开文件目录”按钮，改到文件列表右键菜单：打开文件、打开所在目录、
;           恢复上一个版本（用备份目录里该文件最近一份备份覆盖当前文件，需二次确认）；
;        5. 编译时用 pdf_page_icon.ico 作为 EXE 图标。
;   v1.3 1. 改为“原地修改”：页码直接写回原文件，不再生成“原名_加页码.pdf”；
;           后端先写同目录临时文件，成功后覆盖原文件，失败则删临时文件、原文件不动；
;        2. 修改前自动备份到“我的文档\PDF加页码文件备份\原名_yyyyMMdd_HHmmss.pdf”
;           （窗口里可改目录、可关闭），备份失败就跳过该文件，宁可不动也不留隐患；
;        3. 取消处理完成后的模态汇总框，改为屏幕右下角浮动提示（窗口/拖放/右键三条
;           路径统一），失败时提示框里带前几条原因；--quiet 仍然完全无提示；
;        4. 每个文件的成败追加一行到程序目录的 PDF加页码执行日志.txt（时间/结果/路径/说明）；
;        5. 去掉“文件名以 _加页码 结尾就跳过”的旧规则（原地修改后不会再产生这种文件）。
;   v1.2 1. 编译版改为自带后端：PyInstaller onedir 压成 zip，用 FileInstall 嵌进 EXE；
;           首次运行解压到 %LOCALAPPDATA%\PdfPageNumberTool\backend-1.2\，
;           写完版本标记后不再重复解压，升级版本会自动重解压并清掉旧目录；
;        2. 新增 EnsureBackend() / ExtractZip() / ResolveRunner()：编译版只用内嵌后端，
;           源码态仍用本机 Python，两条路径共用同一个命令行协议，ProcessOne() 只换前缀；
;        3. 编译版不再提示“请安装 Python”，后端缺失/解压失败时提示重新获取 EXE。
;   v1.1 1. 新增图形界面：双击（无参数）先打开窗口，把 PDF 或文件夹拖进列表，
;           点“开始添加页码”才处理；页码参数（起始页、字号、距右边、距下边）
;           可在界面上修改，并保存在 HKCU\Software\PdfPageNumberTool，下次打开沿用；
;        2. 带文件参数启动（拖到 EXE 图标上、资源管理器右键、命令行）仍然不显示
;           窗口，直接走原静默流程，处理完只弹一次汇总；--quiet 行为不变；
;        3. 窗口新增“在 PDF 文件右键菜单中添加「添加页码(&Z)」”选项，勾选/取消立即
;           写入或删除 HKCU 下 .pdf 与 SystemFileAssociations 两处注册表项；编译版
;           启动时若注册的命令路径已失效会自动重写；
;        4. 窗口中拖入含子文件夹的文件夹时询问是否递归扫描（默认“是”），
;           静默模式仍只扫第一层；
;        5. 处理期间锁定界面并逐条更新列表状态，避免看起来像卡死。
;=====================================================================

; ---------- 可调参数 / 默认值：界面上的值会覆盖这里，并写回注册表 ----------
PDF_TOOL_VERSION := "1.12"    ; 工具内部版本号（文件名保持不变，避免 EXE 和右键路径变化）
PDF_START := 1               ; 第一页显示的页码
PDF_FONT_SIZE := 10.5        ; 页码字号（pt），10.5 = 五号
PDF_MARGIN_RIGHT := 2.0      ; 水平边距（cm）：居右=距右边，居左=距左边，居中忽略
PDF_MARGIN_BOTTOM := 1.2     ; 页码基线距页面下边（cm）

; 页码格式列表（v1.4）：文案与公文排版 6.8.19 的 PageFormatList 保持一致
; （公文那边的第一项是“无”，本工具只负责加页码，选了没有意义，所以不列）。
; 顺序固定为 3 种格式 × 左/中/右，解析函数 PdfFormatArgs() 依赖 InStr 判断文案，
; 改文案时必须同步改那个函数。
PDF_FORMAT_DEFAULT := "加总页码 1 / n 居右"
PDF_FORMAT_LIST := [
    "普通数字 1     居左", "普通数字 1     居中", "普通数字 1     居右",
    "格式数字 — 1 — 居左", "格式数字 — 1 — 居中", "格式数字 — 1 — 居右",
    "加总页码 1 / n 居左", "加总页码 1 / n 居中", "加总页码 1 / n 居右"
]
PDF_NUM_FORMAT := PDF_FORMAT_DEFAULT   ; 实际值由 LoadPdfSettings() 从注册表读

; 备份：定位与命名规则与公文排版（6.8.19）保持一致，方便用户在同一种目录结构里找文件
PDF_CREATE_BAK := 1                                        ; 1=修改原文件前先备份
PDF_DEFAULT_BAK_DIR := A_MyDocuments "\PDF加页码文件备份"   ; 默认备份目录
PDF_BAK_DIR := ""                                          ; 实际值由 LoadPdfSettings() 从注册表读
PDF_LOG_NAME := "PDF加页码执行日志.txt"                     ; 追加写入，不轮转

; 独立于公文排版（DocFormatTool）的一套设置，互不影响
PDF_REG_PATH := "HKCU\Software\PdfPageNumberTool"
PDF_CONTEXT_TITLE := "PDF 文件添加页码(&Z)"   ; 右键菜单显示的文字，(&Z) 表示快捷键 Z
PDF_CONTEXT_KEY := "PdfAddPageNum"    ; 右键菜单在注册表里的键名

PdfContextActive := 0        ; 右键菜单是否由本工具注册（与注册表实际状态同步）
PdfRunning := false          ; 是否正在批量处理（处理期间禁止改参数、关闭窗口）
; 必须初始化：AHK v2 里未赋值的变量在表达式中读取会直接报错，
; HidePdfToast 第一次被调用时就会卡在错误对话框里。
PdfToastGui := ""

Main()


;=====================================================================
; [程序入口]
; 关键分流：只有“没有任何文件参数”才建窗口；带参数一律静默处理。
; 这样拖到 EXE 图标上和资源管理器右键调用都不会闪出主窗口。
;=====================================================================
Main() {
    global PdfContextActive

    ; 先摘出 --quiet：命令行批量时用，不弹结果框，改用退出码表示成败（0 全成功 / 1 有失败）
    quiet := false
    rawArgs := []
    for arg in A_Args {
        if (StrLower(Trim(arg, " `t`"")) = "--quiet") {
            quiet := true
            continue
        }
        rawArgs.Push(arg)
    }

    LoadPdfSettings()

    ; 编译版：用户曾勾选过右键菜单，且注册的命令指向的程序已被移动/删除时自动重写。
    ; 未编译时不自动修复，避免把测试用的 .ahk 路径写进用户的右键菜单。
    if (A_IsCompiled && PdfContextActive == 1)
        RepairPdfContextMenu()

    if (rawArgs.Length = 0) {
        ; 无参数（双击）：打开窗口。带 --quiet 说明是命令行/计划任务调用，不能弹窗口
        if (quiet)
            ExitApp 1
        ShowPdfGui()
        return
    }

    ; 带文件参数（拖到 EXE 上、右键菜单、命令行）：不建窗口，直接静默处理
    RunSilent(rawArgs, quiet)
}


;=====================================================================
; [静默流水线] 拖到 EXE 上 / 右键菜单 / 命令行使用，全程不建主窗口
;=====================================================================
RunSilent(rawArgs, quiet) {
    global PDF_START, PDF_FONT_SIZE, PDF_MARGIN_RIGHT, PDF_MARGIN_BOTTOM

    ; 解析后端：编译版用内嵌后端（首次运行会解压），源码态用本机 Python
    runner := ResolveRunner()
    if (runner.exe = "") {
        Notify(runner.err, "PDF 加页码 - 无法启动", "Icon!", quiet)
        return
    }

    ; ---------- 收集要处理的文件 ----------
    targets := []      ; 待处理的 PDF 绝对路径
    results := []      ; 每个文件的结果，最后一次性汇总
    seen := Map()      ; 去重：同一个文件拖两次只处理一次
    skipped := 0       ; 被跳过的重复文件数量

    for arg in rawArgs {
        p := Trim(arg, " `t`"")            ; 拖放有时会带引号
        if (p = "")
            continue
        ; 文件夹末尾的 \ 会让 "\*.pdf" 变成双斜杠，去掉更稳
        if (StrLen(p) > 3 && SubStr(p, -1) = "\")
            p := SubStr(p, 1, -1)

        attr := FileExist(p)
        if (attr = "") {
            results.Push(FailPdfResult(p, "路径不存在"))
            continue
        }

        if InStr(attr, "D") {
            ; 静默模式只扫第一层：公文附件经常是一个文件夹里放一堆 PDF
            Loop Files, p "\*.pdf", "F"
                AddPdfFile(A_LoopFileFullPath, &targets, seen, &skipped)
        } else {
            AddPdfFile(p, &targets, seen, &skipped)
        }
    }

    if (targets.Length = 0 && results.Length = 0) {
        Notify("没有找到可处理的 PDF 文件。`n`n文件夹只扫描第一层，"
            . "请确认里面直接放着 .pdf 文件。", "PDF 加页码", "Icon!", quiet)
        return
    }

    ; Python 输出用 UTF-8，避免中文错误信息变成乱码
    EnvSet "PYTHONUTF8", "1"
    EnvSet "PYTHONIOENCODING", "utf-8"

    ; 大文件叠加要几秒，先把鼠标变成“忙”状态，处理完再恢复
    SetCursorBusy(true)
    idx := 0
    for file in targets {
        idx += 1
        results.Push(ProcessOne(file, idx, runner.exe, runner.script))
    }
    SetCursorBusy(false)

    ; quiet 模式：完全无提示，用退出码告诉调用方成败（0 全部成功 / 1 有失败或出错）
    failed := 0
    for r in results {
        if !r.ok
            failed += 1
    }
    exitCode := failed = 0 ? 0 : 1
    if quiet
        ExitApp exitCode

    ; 非 quiet：右下角提示框停几秒后自动退出，不再弹模态汇总框
    ShowPdfResultAndExit(results, skipped, exitCode)
}


;=====================================================================
; [主窗口]
; 布局与公文排版首页保持同一风格：顶部标题、参数组框、任务列表、选项、底部按钮。
; 窗口只由 Main() 的无参数分支调用，静默流程永远不会走到这里。
;=====================================================================
ShowPdfGui() {
    global PDF_TOOL_VERSION, PDF_START, PDF_FONT_SIZE, PDF_MARGIN_RIGHT, PDF_MARGIN_BOTTOM
    global PDF_CREATE_BAK, PDF_BAK_DIR, PDF_NUM_FORMAT, PDF_FORMAT_LIST
    global pdfGui, uiPdfStart, uiPdfFont, uiPdfRight, uiPdfBottom, uiPdfFormat
    global uiPdfLV, uiPdfCtx, uiPdfSB, uiPdfRun, uiPdfClear, uiPdfRemove
    global uiPdfBak, uiPdfBakDir, uiPdfBakBtn, uiPdfBakOpen
    global PdfContextActive

    ; 已经建过窗口时（例如从托盘再次打开）只显示，不重复创建
    if (IsSet(pdfGui) && IsObject(pdfGui)) {
        try {
            pdfGui.Show()
            return
        }
    }

    pdfGui := Gui("+Theme -MaximizeBox", "PDF 加页码工具 V" PDF_TOOL_VERSION)
    pdfGui.OnEvent("Close", PdfGuiClose)
    pdfGui.OnEvent("DropFiles", PdfGuiDropFiles)
    pdfGui.SetFont("s10", "Microsoft YaHei")
    pdfGui.BackColor := "FFFFFF"

    ; --- 顶部横幅 ---
    pdfGui.SetFont("s18 bold c0055AA")
    pdfGui.Add("Text", "x20 y15 w660 Center BackgroundTrans", "PDF 加页码工具")
    pdfGui.SetFont("s9 norm c666666")
    pdfGui.Add("Text", "x20 y45 w660 Center BackgroundTrans",
        "页码叠加在页脚位置（格式可选）  |  原 PDF 内容不改动  |  把 PDF 或文件夹拖入下方列表")
    pdfGui.SetFont("s10 c000000")

    ; --- 页码参数 ---
    pdfGui.Add("GroupBox", "x20 y75 w660 h140", "页码参数")
    pdfGui.Add("Text", "x35 y102 w85 Right", "起始页码")
    uiPdfStart := pdfGui.Add("Edit", "x125 y98 w55")
    uiPdfStart.Value := PDF_START

    pdfGui.Add("Text", "x195 y102 w85 Right", "字号(pt)")
    uiPdfFont := pdfGui.Add("Edit", "x285 y98 w55")
    uiPdfFont.Value := PDF_FONT_SIZE

    pdfGui.Add("Text", "x355 y102 w85 Right", "水平边距(cm)")
    uiPdfRight := pdfGui.Add("Edit", "x445 y98 w55")
    uiPdfRight.Value := PDF_MARGIN_RIGHT

    pdfGui.Add("Text", "x515 y102 w85 Right", "距下边(cm)")
    uiPdfBottom := pdfGui.Add("Edit", "x605 y98 w55")
    uiPdfBottom.Value := PDF_MARGIN_BOTTOM

    ; 页码格式：文案与公文排版一致；Choose 用注册表里保存的文案定位，
    ; 找不到（旧版本/手工改过注册表）时回落到第一项“普通数字居左”，
    ; 真正的回落默认值在 LoadPdfSettings() 里已归一为 PDF_FORMAT_DEFAULT。
    pdfGui.Add("Text", "x35 y138 w85 Right", "页码格式")
    formatIdx := 1
    for idx, item in PDF_FORMAT_LIST {
        if (item = PDF_NUM_FORMAT) {
            formatIdx := idx
            break
        }
    }
    uiPdfFormat := pdfGui.Add("DropDownList", "x125 y134 w540 Choose" formatIdx, PDF_FORMAT_LIST)
    uiPdfStart.OnEvent("Change", (ctrl, *) => SavePdfParamNow("StartNum", ctrl.Value))
    uiPdfFont.OnEvent("Change", (ctrl, *) => SavePdfParamNow("FontSize", ctrl.Value))
    uiPdfRight.OnEvent("Change", (ctrl, *) => SavePdfParamNow("MarginRight", ctrl.Value))
    uiPdfBottom.OnEvent("Change", (ctrl, *) => SavePdfParamNow("MarginBottom", ctrl.Value))
    uiPdfFormat.OnEvent("Change", (ctrl, *) => SavePdfParamNow("PageNumFormat", ctrl.Text))

    pdfGui.SetFont("s9 c808080")
    pdfGui.Add("Text", "x35 y166 w630",
        "输出：页码直接写回原文件（不再生成 _加页码.pdf）；修改前自动备份，备份失败就不动原文件。")
    pdfGui.Add("Text", "x35 y190 w630",
        "页码以透明图层叠在最上层，扫描件也能加；水平边距在“居中”时忽略。")
    pdfGui.SetFont("s10 c000000")

    ; --- 选项：备份 + 右键菜单 ---
    pdfGui.Add("GroupBox", "x20 y225 w660 h155", "选项")
    uiPdfBak := pdfGui.Add("Checkbox", "x35 y250 w300", "修改原文件前先备份")
    uiPdfBak.Value := PDF_CREATE_BAK            ; 1=勾选，0=不勾选
    uiPdfBak.OnEvent("Click", OnPdfBakToggle)

    pdfGui.Add("Text", "x35 y282 w65", "备份目录")
    uiPdfBakDir := pdfGui.Add("Edit", "x105 y278 w430", PDF_BAK_DIR)
    ; 手工改路径时立即保存，下次打开沿用（路径不合法时会在处理前再校验一次）
    uiPdfBakDir.OnEvent("Change", (ctrl, *) => SavePdfBakDir(ctrl.Value))
    uiPdfBakBtn := pdfGui.Add("Button", "x545 y276 w80 h40", "选择…")
    uiPdfBakBtn.OnEvent("Click", SelectPdfBakDir)
    uiPdfBakOpen := pdfGui.Add("Button", "x635 y276 w45 h40", "打开")
    uiPdfBakOpen.OnEvent("Click", OpenPdfBakFolder)

    pdfGui.SetFont("s8 c909090")
    pdfGui.Add("Text", "x35 y307 w630",
        "备份文件名：原名_年月日_时分秒.pdf，每次处理单独留一份；目录不存在会自动创建。")
    pdfGui.SetFont("s10 c000000")

    uiPdfCtx := pdfGui.Add("Checkbox", "x35 y343 w620",
        "在 PDF 文件右键菜单中添加「" PDF_CONTEXT_TITLE "」")
    uiPdfCtx.OnEvent("Click", OnPdfContextToggle)
    pdfGui.SetFont("s8 c909090")
    pdfGui.Add("Text", "x35 y366 w620",
        "勾选后可在资源管理器右键 PDF 直接加页码（静默处理，不显示本窗口）；取消勾选即删除该菜单项。")
    pdfGui.SetFont("s10 c000000")

    ; --- 任务列表 ---
    pdfGui.Add("GroupBox", "x20 y390 w660 h215", "PDF 任务列表 (拖拽 PDF 文件、文件夹至此)")
    uiPdfLV := pdfGui.Add("ListView", "x30 y416 w640 h178 Grid", ["文件名", "状态", "文件路径"])
    uiPdfLV.ModifyCol(1, 200)
    uiPdfLV.ModifyCol(2, 120)
    uiPdfLV.ModifyCol(3, 300)

    ; 列表右键菜单：只对选中的行有效，避免误操作。
    ; “恢复上一个版本”会用备份覆盖原文件，属于破坏性操作，函数里还会再弹一次确认。
    lvMenu := Menu()
    lvMenu.Add("打开文件", (*) => OpenSelectedPdfFile())
    lvMenu.Add("打开所在目录", (*) => OpenSelectedPdfFolder())
    lvMenu.Add()                       ; 分隔线
    lvMenu.Add("恢复上一个版本", (*) => RestoreSelectedPdfBackup())
    lvMenu.Add()
    lvMenu.Add("移除所选文件", (*) => RemoveSelectedPdfRows())
    uiPdfLV.OnEvent("ContextMenu", (ctrl, item, isRightClick, x, y) => (item > 0 ? lvMenu.Show() : ""))
    uiPdfLV.OnEvent("DoubleClick", (*) => OpenSelectedPdfFile())

    ; --- 底部按钮 ---
    ; v1.4 去掉“打开文件目录”（已移到列表右键菜单），三个按钮重新分配宽度
    uiPdfClear := pdfGui.Add("Button", "x20 y615 w140 h35", "清空列表")
    uiPdfClear.OnEvent("Click", ClearPdfList)

    uiPdfRemove := pdfGui.Add("Button", "x175 y615 w140 h35", "移除所选")
    uiPdfRemove.OnEvent("Click", (*) => RemoveSelectedPdfRows())

    ; 运行按钮用浅绿色背景，和公文排版的“保存参数并运行”保持一致
    uiPdfRun := pdfGui.Add("Text", "x340 y615 w340 h35 Center +0x200 BackgroundE8F5E9 Border c117711",
        "开始添加页码  ▶")
    uiPdfRun.SetFont("s11 bold")
    uiPdfRun.OnEvent("Click", RunGuiProcessing)

    uiPdfSB := pdfGui.Add("StatusBar",, " 准备就绪：把 PDF 文件或文件夹拖入上方列表。")
    pdfGui.Show("w700 h685")

    ; 勾选状态以注册表实际存在为准：用户手工删过菜单时，配置里的旧值不会被误显示成“已勾选”
    uiPdfCtx.Value := PdfContextMenuExists() ? 1 : 0
    if (uiPdfCtx.Value != PdfContextActive) {
        PdfContextActive := uiPdfCtx.Value
        SetPdfVal("ContextActive", PdfContextActive)
    }
}

; 关闭窗口：批量处理期间不允许直接关掉，避免只处理了一半就退出
PdfGuiClose(*) {
    global PdfRunning

    if (PdfRunning) {
        ShowPdfToast("正在处理", "请等当前批次处理完成后再关闭窗口。", false)
        return true   ; 返回非空值阻止窗口销毁
    }
    ExitApp()
}


;=====================================================================
; [窗口：拖入文件]
; 只收 .pdf；文件夹默认只扫第一层，若含子文件夹则先询问（与公文排版一致）。
;=====================================================================
PdfGuiDropFiles(guiObj, ctrlObj, fileArray, x, y) {
    global uiPdfLV, uiPdfSB, PdfRunning

    if (PdfRunning) {
        uiPdfSB.SetText(" 正在处理，请等当前批次结束后再添加文件。")
        return
    }

    ; 拖入的文件夹若包含子文件夹，先询问是否连同子文件夹一起扫描；默认按钮为“是”
    scanSubfolders := false
    for path in fileArray {
        if (InStr(FileExist(path), "D") && HasSubFolders(path)) {
            answer := MsgBox("检测到拖入的文件夹中包含子文件夹。`n`n"
                . "是否扫描子文件夹中的 PDF 文件？`n`n"
                . "· 选择“是”：包含所有子文件夹中的文件（推荐）`n"
                . "· 选择“否”：只添加该文件夹第一层的文件", "扫描子文件夹", "YesNo Icon?")
            scanSubfolders := (answer == "Yes")
            break
        }
    }

    ; 列表里已有的路径先放进 Map，同一个文件拖两次只保留一条
    seen := Map()
    Loop uiPdfLV.GetCount() {
        p := uiPdfLV.GetText(A_Index, 3)
        if (p != "")
            seen[StrLower(p)] := true
    }

    added := 0, skipped := 0
    for path in fileArray {
        if InStr(FileExist(path), "D") {
            ; FR = 文件 + 递归子文件夹；F = 仅当前文件夹第一层
            fileOpts := scanSubfolders ? "FR" : "F"
            Loop Files, path "\*.pdf", fileOpts
                PdfAddToGui(A_LoopFileFullPath, &seen, &added, &skipped)
        } else {
            PdfAddToGui(path, &seen, &added, &skipped)
        }
    }

    msg := " 已添加 " added " 个 PDF 任务。"
    if (skipped > 0)
        msg .= " 跳过 " skipped " 个（重复或锁文件）。"
    msg .= " 当前列表共 " uiPdfLV.GetCount() " 个。"
    uiPdfSB.SetText(msg)
}

; 把单个文件加入窗口列表（过滤非 PDF、Office 锁文件）
; v1.3 起原地修改，不会再产生 _加页码.pdf，所以不再按后缀跳过
PdfAddToGui(path, &seen, &added, &skipped) {
    global uiPdfLV

    SplitFilePath(path, &dir, &name, &ext)
    if (StrLower(ext) != "pdf")
        return                                   ; 文件夹里混着其他文件很正常，静默跳过
    if (SubStr(name, 1, 2) = "~$")
        return                                   ; Word/WPS 打开文档时生成的锁文件

    key := StrLower(path)
    if (seen.Has(key)) {
        skipped += 1
        return
    }
    seen[key] := true

    SplitPath path, &fileName
    uiPdfLV.Add("", fileName, "等待处理", path)
    added += 1
}


;=====================================================================
; [窗口：开始处理]
; 先把输入框里的参数校验一遍（输入框允许手工输入，不能把 "abc" 传给 Python），
; 再逐个文件处理。每个文件处理完更新一次列表状态，然后交给定时器处理下一个，
; 这样窗口才有机会重绘，不会看起来像卡死。
;=====================================================================
RunGuiProcessing(*) {
    global uiPdfStart, uiPdfFont, uiPdfRight, uiPdfBottom, uiPdfFormat
    global uiPdfLV, uiPdfSB, uiPdfCtx, uiPdfRun, uiPdfClear, uiPdfRemove
    global PDF_START, PDF_FONT_SIZE, PDF_MARGIN_RIGHT, PDF_MARGIN_BOTTOM, PDF_NUM_FORMAT
    global PdfRunning, PdfRunQueue, PdfRunIndex, PdfRunResults, PdfRunExe, PdfRunScript

    if (PdfRunning)
        return

    ; v1.6 起不再弹模态框：状态栏留一行原因，同时右下角浮一个提示框
    if (uiPdfLV.GetCount() = 0) {
        uiPdfSB.SetText(" 列表里还没有 PDF，请先把 PDF 文件或文件夹拖进列表。")
        ShowPdfToast("列表里还没有 PDF", "请先把 PDF 文件或文件夹拖进列表。", false, 3600)
        return
    }

    startText := Trim(uiPdfStart.Value)
    fontText := Trim(uiPdfFont.Value)
    rightText := Trim(uiPdfRight.Value)
    bottomText := Trim(uiPdfBottom.Value)
    formatText := uiPdfFormat.Text

    if (!IsNumber(startText) || Integer(startText) < 1 || Integer(startText) > 999999) {
        uiPdfSB.SetText(" 参数有误：起始页码必须是 1 到 999999 之间的整数。")
        ShowPdfToast("参数有误", "起始页码必须是 1 到 999999 之间的整数。", false, 3600)
        uiPdfStart.Focus()
        return
    }
    if (!IsNumber(fontText) || Float(fontText) < 5 || Float(fontText) > 72) {
        uiPdfSB.SetText(" 参数有误：页码字号必须在 5 到 72 磅之间。")
        ShowPdfToast("参数有误", "页码字号必须在 5 到 72 磅之间。", false, 3600)
        uiPdfFont.Focus()
        return
    }
    if (!IsNumber(rightText) || Float(rightText) < 0 || Float(rightText) > 20) {
        uiPdfSB.SetText(" 参数有误：水平边距必须在 0 到 20 厘米之间。")
        ShowPdfToast("参数有误", "水平边距必须在 0 到 20 厘米之间。", false, 3600)
        uiPdfRight.Focus()
        return
    }
    if (!IsNumber(bottomText) || Float(bottomText) < 0 || Float(bottomText) > 20) {
        uiPdfSB.SetText(" 参数有误：距下边必须在 0 到 20 厘米之间。")
        ShowPdfToast("参数有误", "距下边必须在 0 到 20 厘米之间。", false, 3600)
        uiPdfBottom.Focus()
        return
    }

    ; 参数校验通过后才写回全局变量和注册表，界面显示与 Python 实际使用保持一致
    PDF_START := Integer(startText)
    PDF_FONT_SIZE := Float(fontText)
    PDF_MARGIN_RIGHT := Float(rightText)
    PDF_MARGIN_BOTTOM := Float(bottomText)
    PDF_NUM_FORMAT := formatText
    SetPdfVal("StartNum", PDF_START)
    SetPdfVal("FontSize", PDF_FONT_SIZE)
    SetPdfVal("MarginRight", PDF_MARGIN_RIGHT)
    SetPdfVal("MarginBottom", PDF_MARGIN_BOTTOM)
    SetPdfVal("PageNumFormat", PDF_NUM_FORMAT)

    ; 解析后端（编译版首次运行会先解压内嵌后端，约 2–3 秒，只发生一次）
    runner := ResolveRunner()
    if (runner.exe = "") {
        ; 状态栏只有一行，这里取错误文本的第一行；完整原因（多行）在提示框里，日志也会记
        firstErr := StrSplit(runner.err, "`n", "`r")[1]
        uiPdfSB.SetText(" 无法启动后端：" firstErr)
        ShowPdfToast("无法启动后端", runner.err, false, 6000)
        return
    }

    EnvSet "PYTHONUTF8", "1"
    EnvSet "PYTHONIOENCODING", "utf-8"

    ; 列表里的行号在处理过程中不会变（期间禁止增删），所以直接记下行号用于回写状态
    PdfRunQueue := []
    Loop uiPdfLV.GetCount() {
        row := A_Index
        p := uiPdfLV.GetText(row, 3)
        if (p = "")
            continue
        PdfRunQueue.Push({ row: row, path: p, name: uiPdfLV.GetText(row, 1) })
        SetPdfRowStatus(row, "等待处理")
    }

    if (PdfRunQueue.Length = 0) {
        uiPdfSB.SetText(" 列表里没有有效的 PDF 路径。")
        ShowPdfToast("没有可处理的 PDF", "列表里没有有效的 PDF 路径。", false, 3600)
        return
    }

    PdfRunIndex := 0
    PdfRunResults := []
    PdfRunExe := runner.exe
    PdfRunScript := runner.script
    PdfRunning := true

    SetPdfControlsLocked(true)
    SetCursorBusy(true)
    uiPdfSB.SetText(" 正在准备处理 " PdfRunQueue.Length " 个文件…")

    ; -10 表示只执行一次；每个文件处理完再续一次，让消息循环有机会重绘界面
    SetTimer(PdfProcessNext, -10)
}

; 处理队列里的下一个文件；全部处理完后收尾
PdfProcessNext() {
    global uiPdfLV, uiPdfSB, PdfRunning
    global PdfRunQueue, PdfRunIndex, PdfRunResults, PdfRunExe, PdfRunScript

    if (!PdfRunning)
        return

    PdfRunIndex += 1
    total := PdfRunQueue.Length
    if (PdfRunIndex > total) {
        FinishGuiRun()
        return
    }

    item := PdfRunQueue[PdfRunIndex]
    uiPdfLV.Modify(item.row, "", item.name, "处理中…")
    uiPdfSB.SetText(" 正在处理 " PdfRunIndex "/" total "：" item.name)

    ; 下面 RunWait 会阻塞脚本线程，消息循环停止，界面不会自动重绘。
    ; 这里强制把主窗口和所有子控件立即重画一次，用户才能看到“处理中 3/5”这类进度。
    ForceGuiRepaint()

    ; RunWait 会阻塞当前线程，但每个文件之间会回到消息循环，界面状态能及时刷新
    r := ProcessOne(item.path, PdfRunIndex, PdfRunExe, PdfRunScript)
    PdfRunResults.Push(r)
    SetPdfRowStatus(item.row, r.ok ? "成功" : "失败")

    SetTimer(PdfProcessNext, -10)
}

; 立即重绘主窗口及所有子控件（RunWait 阻塞期间唯一的刷新机会）
ForceGuiRepaint() {
    global pdfGui

    if (IsObject(pdfGui)) {
        ; RDW_INVALIDATE(0x0001) | RDW_UPDATENOW(0x0100) | RDW_ALLCHILDREN(0x0080)
        DllCall("RedrawWindow", "Ptr", pdfGui.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0181)
    }
}

; 本批次结束：恢复界面、写状态栏，右下角弹提示框（不再弹模态汇总）
FinishGuiRun() {
    global uiPdfSB, PdfRunning, PdfRunResults

    PdfRunning := false
    SetPdfControlsLocked(false)
    SetCursorBusy(false)

    okCount := 0
    for r in PdfRunResults {
        if (r.ok)
            okCount += 1
    }
    uiPdfSB.SetText(" 处理完成：成功 " okCount " 个，失败 " (PdfRunResults.Length - okCount) " 个。")
    ShowPdfResult(PdfRunResults)
}

; 只改列表第 2 列（状态），第 1 列文字原样传回，避免清空文件名
SetPdfRowStatus(row, status) {
    global uiPdfLV

    uiPdfLV.Modify(row, "", uiPdfLV.GetText(row, 1), status)
}

; 批量处理期间禁用所有会改变任务或参数的控件；
; 列表本身保持可用（可以滚动查看进度），但增删入口都已禁用
SetPdfControlsLocked(locked) {
    global uiPdfStart, uiPdfFont, uiPdfRight, uiPdfBottom, uiPdfFormat
    global uiPdfCtx, uiPdfRun, uiPdfClear, uiPdfRemove
    global uiPdfBak, uiPdfBakDir, uiPdfBakBtn, uiPdfBakOpen

    enabled := !locked
    for ctrl in [uiPdfStart, uiPdfFont, uiPdfRight, uiPdfBottom, uiPdfFormat, uiPdfBak,
                 uiPdfBakDir, uiPdfBakBtn, uiPdfBakOpen, uiPdfCtx, uiPdfRun,
                 uiPdfClear, uiPdfRemove]
        ctrl.Enabled := enabled
}


;=====================================================================
; [窗口：列表操作与右键菜单选项]
;=====================================================================
ClearPdfList(*) {
    global uiPdfLV, uiPdfSB, PdfRunning

    if (PdfRunning)
        return
    uiPdfLV.Delete()
    uiPdfSB.SetText(" 列表已清空。")
}

RemoveSelectedPdfRows(*) {
    global uiPdfLV, uiPdfSB, PdfRunning

    if (PdfRunning)
        return

    ; 先收集所有选中行，再从后往前删：从前往后删会让后面的行号整体上移，漏删或删错
    rows := []
    row := 0
    while (row := uiPdfLV.GetNext(row))
        rows.Push(row)

    if (rows.Length = 0) {
        uiPdfSB.SetText(" 请先在列表里选中要移除的文件。")
        return
    }
    while (rows.Length)
        uiPdfLV.Delete(rows.Pop())

    uiPdfSB.SetText(" 已移除所选任务，当前列表共 " uiPdfLV.GetCount() " 个。")
}

OpenSelectedPdfFile(*) {
    global uiPdfLV, uiPdfSB

    row := uiPdfLV.GetNext()
    if (row = 0) {
        uiPdfSB.SetText(" 请先在列表里选中一个文件。")
        return
    }
    path := uiPdfLV.GetText(row, 3)
    try Run('"' path '"')
    catch
        uiPdfSB.SetText(" 无法打开文件：" path)
}

OpenSelectedPdfFolder(*) {
    global uiPdfLV, uiPdfSB

    row := uiPdfLV.GetNext()
    if (row = 0) {
        uiPdfSB.SetText(" 请先在列表里选中一个文件。")
        return
    }
    path := uiPdfLV.GetText(row, 3)
    SplitPath(path, , &dir)
    if (dir = "" || !DirExist(dir)) {
        uiPdfSB.SetText(" 找不到文件所在目录。")
        return
    }
    ; /select 会打开目录并选中该文件；文件刚好被移走时退回只打开目录
    try Run('explorer.exe /select,"' path '"')
    catch
        try Run('"' dir '"')
}

;---------------------------------------------------------------------
; v1.8：新增“打开”按钮（OpenPdfBakFolder）的逻辑：打开当前备份目录。
; 与主工具 6.8.19.ahk 的 OpenBakFolder 保持一致：取 Edit 里当前备份目录
; （缺省走 PDF_BAK_DIR / PDF_DEFAULT_BAK_DIR），目录不存在则自动创建，
; 再用 explorer/Run 打开。只做“打开目录”，不做选择/保存。
;---------------------------------------------------------------------
OpenPdfBakFolder(*) {
    global uiPdfBakDir, PDF_BAK_DIR, PDF_DEFAULT_BAK_DIR, uiPdfSB

    dir := Trim(uiPdfBakDir.Value)
    if (dir = "")
        dir := PDF_BAK_DIR
    if (dir = "")
        dir := PDF_DEFAULT_BAK_DIR
    if (!DirExist(dir)) {
        try DirCreate(dir)
        catch {
            uiPdfSB.SetText(" 无法打开备份目录：" dir)
            return
        }
    }
    try Run('"' dir '"')
    catch {
        uiPdfSB.SetText(" 无法打开备份目录，请检查路径或权限！")
    }
}

;---------------------------------------------------------------------
; 恢复上一个版本：用备份目录里该文件最近的一份备份覆盖当前文件。
;
; 这是破坏性操作（当前内容无法撤销），所以顺序固定为：
;   找备份 → 二次确认 → 检查可写 → 整文件覆盖。
; 备份目录同时查“当前设置”和“默认目录”：用户可能中途改过备份路径，
; 旧备份还留在原来的地方，只查一个目录会漏掉。
;
; 备份文件名格式（BackupPdfFile 写入）：原名_yyyyMMdd_HHmmss[ (n)].pdf。
; 这里不用正则硬套整串，只校验时间戳结构，避免文件名里的特殊字符出问题。
;---------------------------------------------------------------------
RestoreSelectedPdfBackup(*) {
    global uiPdfLV, uiPdfSB, PdfRunning
    global PDF_BAK_DIR, PDF_DEFAULT_BAK_DIR

    if (PdfRunning) {
        uiPdfSB.SetText(" 正在处理，请等当前批次结束后再恢复。")
        return
    }

    row := uiPdfLV.GetNext()
    if (row = 0) {
        uiPdfSB.SetText(" 请先在列表里选中一个文件。")
        return
    }
    path := uiPdfLV.GetText(row, 3)
    if !FileExist(path) {
        uiPdfSB.SetText(" 文件已不存在：" path)
        return
    }

    SplitFilePath(path, &dir, &name, &ext)
    prefix := name "_"
    latest := "", latestTime := ""
    seenDirs := Map()
    for bakDir in [Trim(PDF_BAK_DIR, " `t`""), PDF_DEFAULT_BAK_DIR] {
        if (bakDir = "" || seenDirs.Has(StrLower(bakDir)))
            continue
        seenDirs[StrLower(bakDir)] := true
        if !DirExist(bakDir)
            continue
        Loop Files, bakDir "\" name "_*." ext, "F" {
            ; 去掉“原名_”前缀和扩展名，剩下必须是 yyyyMMdd_HHmmss 或再带 (2)
            tail := SubStr(A_LoopFileName, StrLen(prefix) + 1)
            if (StrLen(tail) < StrLen(ext) + 1 || StrLower(SubStr(tail, -StrLen(ext))) != "." StrLower(ext))
                continue
            stamp := SubStr(tail, 1, StrLen(tail) - StrLen(ext) - 1)
            if (StrLen(stamp) < 15 || SubStr(stamp, 9, 1) != "_")
                continue
            if (!IsNumber(SubStr(stamp, 1, 8)) || !IsNumber(SubStr(stamp, 10, 6)))
                continue
            rest := SubStr(stamp, 16)
            if (rest != "" && !RegExMatch(rest, "^\(\d+\)$"))
                continue

            ; 同一秒的多份备份用修改时间排先后，时间相同再按文件名比大小
            t := A_LoopFileTimeModified
            if (latest = "" || t > latestTime || (t = latestTime && A_LoopFileName > latestName)) {
                latestTime := t
                latestName := A_LoopFileName
                latest := A_LoopFileFullPath
            }
        }
    }

    if (latest = "") {
        uiPdfSB.SetText(" 备份目录里没有找到该文件的历史备份。")
        ShowPdfToast("没有可恢复的备份", "备份目录里没找到 " name "." ext " 的历史备份。", false, 3200)
        return
    }

    ; 默认按钮是“否”，避免用户顺手回车就把当前文件覆盖掉
    answer := MsgBox("将用最近一份备份覆盖当前文件：`n`n当前：" path "`n备份：" latest "`n`n当前内容会被替换且无法撤销，是否继续？",
        "恢复上一个版本", "YesNo Icon? Default2")
    if (answer != "Yes")
        return

    if !IsPdfWritable(path) {
        uiPdfSB.SetText(" 文件被占用或没有写入权限，无法恢复。")
        ShowPdfToast("恢复失败", "文件被占用或没有写入权限，请先关闭 PDF 阅读器。", false, 3600)
        return
    }

    ; v1.7：恢复覆盖本身不可撤销，先给当前文件留一份快照（_before-restore_<stamp>），
    ; 快照失败则中止恢复，绝不直接覆盖。快照放备份目录，文件名用“原名_before-restore_…”，
    ; 不会被恢复逻辑对 yyyyMMdd_HHmmss 的结构校验误认为历史备份。
    snapDir := Trim(PDF_BAK_DIR, " `t`"")
    if (snapDir = "")
        snapDir := PDF_DEFAULT_BAK_DIR
    if (!DirExist(snapDir)) {
        try DirCreate snapDir
    }
    SplitFilePath(path, &snapDir2, &snapName, &snapExt)
    snapStamp := FormatTime(A_Now, "yyyyMMdd_HHmmss")
    snapPath := snapDir "\" snapName "_before-restore_" snapStamp "." snapExt
    okSnap := true
    try FileCopy path, snapPath
    catch as errSnap {
        okSnap := false
        ShowPdfToast("恢复失败", "无法备份当前文件，已中止恢复：" errSnap.Message, false, 3600)
    }
    if (!okSnap) {
        uiPdfSB.SetText(" 已中止恢复：备份当前文件失败。")
        return
    }

    try FileCopy latest, path, true
    catch as err {
        uiPdfSB.SetText(" 恢复失败：" err.Message)
        ShowPdfToast("恢复失败", "无法覆盖原文件：" err.Message, false, 3800)
        return
    }

    SetPdfRowStatus(row, "已恢复")
    uiPdfSB.SetText(" 已恢复：" path)
    ShowPdfToast("恢复成功", "已用备份覆盖当前文件：`n" latest, true, 3200)
}

; 勾选/取消“在 PDF 右键菜单中添加…”：立即写注册表，不等点运行
OnPdfContextToggle(ctrl, *) {
    global PdfContextActive, PdfRunning

    if (PdfRunning) {
        ctrl.Value := PdfContextActive ? 1 : 0
        return
    }

    if (ctrl.Value) {
        if (!RegisterPdfContextMenu(true)) {
            ; 部分键写失败时清理掉残留，避免出现“勾着但右侧菜单不完整”的状态
            RemovePdfContextMenu(false)
            ctrl.Value := 0
        }
    } else {
        RemovePdfContextMenu(true)
    }
}


;=====================================================================
; [窗口：备份设置]
; 备份是“原地修改”的唯一回退手段，所以默认开启；关闭时立即用提示条
; 说明风险（不弹确认框，避免打断操作）。
;=====================================================================
OnPdfBakToggle(ctrl, *) {
    global PDF_CREATE_BAK, PdfRunning

    if (PdfRunning) {
        ctrl.Value := PDF_CREATE_BAK
        return
    }

    PDF_CREATE_BAK := ctrl.Value ? 1 : 0
    SetPdfVal("CreateBAK", PDF_CREATE_BAK)
    ShowPdfToast(PDF_CREATE_BAK ? "已开启备份" : "已关闭备份",
        PDF_CREATE_BAK ? "修改原文件前会先备份到下面的目录。" : "注意：关闭后页码会直接写入原文件，无法撤销。",
        PDF_CREATE_BAK, 2600)
}

; 手工编辑备份目录：只在内容变化时保存，避免每次敲键都写注册表
SavePdfBakDir(value) {
    global PDF_BAK_DIR

    dir := Trim(value, " `t`"")
    if (dir = PDF_BAK_DIR)
        return
    PDF_BAK_DIR := dir
    SetPdfVal("BakDir", dir)
}

; 选择备份目录：选完立即写回输入框并保存
SelectPdfBakDir(*) {
    global uiPdfBakDir, PDF_BAK_DIR

    startDir := DirExist(PDF_BAK_DIR) ? PDF_BAK_DIR : A_MyDocuments
    try {
        ; 与公文排版一致：前缀 * 表示“以此目录为起始位置”，3 = 新版选择框
        selected := DirSelect(startDir = "" ? "" : "*" startDir, 3, "请选择备份存放文件夹")
        if (selected = "")
            return
        uiPdfBakDir.Value := selected
        SavePdfBakDir(selected)
    } catch as err {
        ShowPdfToast("选择失败", err.Message, false, 3200)
    }
}


;=====================================================================
; [PDF 文件右键菜单]
; 注册在 .pdf 和 SystemFileAssociations 两处，Explorer 命中任意一处都能显示。
; 命令带 "%1"，所以右键调用属于“带文件参数启动”，不会弹出主窗口。
;=====================================================================
GetPdfContextBasePaths() {
    return [
        "HKCU\Software\Classes\.pdf\shell",
        "HKCU\Software\Classes\SystemFileAssociations\.pdf\Shell"
    ]
}

RegisterPdfContextMenu(showToast := true) {
    global PDF_CONTEXT_TITLE, PDF_CONTEXT_KEY, PdfContextActive

    ; 路径优先级：编译版自身 → 同目录 EXE（未编译测试时也能指向成品）→ AHK + 脚本
    if (A_IsCompiled) {
        cmdPrefix := '"' A_ScriptFullPath '"'
        iconPath := A_ScriptFullPath
    } else if FileExist(A_ScriptDir "\PDF 加页码工具.exe") {
        defaultExe := A_ScriptDir "\PDF 加页码工具.exe"
        cmdPrefix := '"' defaultExe '"'
        iconPath := defaultExe
    } else {
        cmdPrefix := '"' A_AhkPath '" "' A_ScriptFullPath '"'
        iconPath := A_AhkPath
    }

    hasError := false
    for _, basePath in GetPdfContextBasePaths() {
        try {
            runKey := basePath "\" PDF_CONTEXT_KEY
            RegCreateKey(runKey "\command")
            RegWrite(PDF_CONTEXT_TITLE, "REG_SZ", runKey)
            RegWrite(PDF_CONTEXT_TITLE, "REG_SZ", runKey, "MUIVerb")
            RegWrite(iconPath, "REG_SZ", runKey, "Icon")
            RegWrite(cmdPrefix ' "%1"', "REG_SZ", runKey "\command")
        } catch {
            hasError := true
        }
    }

    if (!hasError) {
        PdfContextActive := 1
        SetPdfVal("ContextActive", 1)
        if (showToast)
            ShowPdfToast("右键菜单已添加", "在资源管理器右键 PDF 即可看到「" PDF_CONTEXT_TITLE "」", true)
    } else {
        if (showToast)
            ShowPdfToast("添加失败", "部分注册表项写入受阻，请检查权限后重试。", false)
    }
    return !hasError
}

RemovePdfContextMenu(showToast := true) {
    global PDF_CONTEXT_KEY, PdfContextActive

    for _, basePath in GetPdfContextBasePaths() {
        try RegDeleteKey(basePath "\" PDF_CONTEXT_KEY)
    }

    PdfContextActive := 0
    SetPdfVal("ContextActive", 0)
    if (showToast)
        ShowPdfToast("右键菜单已删除", "资源管理器里 PDF 的「添加页码」菜单已移除。", true)
    return true
}

; 判断右键菜单是否真的存在：以注册表为准，不看配置里的旧值
PdfContextMenuExists() {
    global PDF_CONTEXT_KEY

    for _, basePath in GetPdfContextBasePaths() {
        try {
            RegRead(basePath "\" PDF_CONTEXT_KEY "\command", "")
            return true
        }
    }
    return false
}

; 编译版启动时自愈：命令里指向的程序被移动/删除后重新写入正确路径
RepairPdfContextMenu() {
    global PDF_CONTEXT_KEY, PDF_CONTEXT_TITLE

    keyPath := "HKCU\Software\Classes\.pdf\shell\" PDF_CONTEXT_KEY "\command"
    try {
        cmd := RegRead(keyPath, "")
    } catch {
        ; 键不存在（用户只注册了另一处，或手工删过）：尊重现状，不自动重建
        return
    }

    ; 命令形如 "C:\...\PDF 加页码工具.exe" "%1"，取出第一个引号内的路径
    exePath := ""
    if RegExMatch(cmd, '^\s*"([^"]+)"', &m)
        exePath := m[1]

    ; v1.4 改了菜单文字（PDF 文件添加页码）。已注册旧文案的用户不会主动
    ; 重新勾选，所以启动时顺便比对一次，文字不一致就整体重写。
    title := ""
    try title := RegRead("HKCU\Software\Classes\.pdf\shell\" PDF_CONTEXT_KEY, "")

    if ((exePath != "" && !FileExist(exePath)) || title != PDF_CONTEXT_TITLE)
        RegisterPdfContextMenu(false)
}


; 把下拉框文案翻译成后端参数。必须和 PDF_FORMAT_LIST 的文案同步：
;   含“格式数字”→ dash，含“加总页码”→ total，其余（普通数字）→ plain；
;   含“居左”→ left，含“居右”→ right，其余（居中）→ center。
PdfFormatArgs(formatStr) {
    style := "plain"
    if InStr(formatStr, "格式数字")
        style := "dash"
    else if InStr(formatStr, "加总页码")
        style := "total"

    align := "center"
    if InStr(formatStr, "居左")
        align := "left"
    else if InStr(formatStr, "居右")
        align := "right"

    return { style: style, align: align }
}


;=====================================================================
; [设置读写] 独立注册表键，避免和公文排版（DocFormatTool）混在一起
;=====================================================================
LoadPdfSettings() {
    global PDF_START, PDF_FONT_SIZE, PDF_MARGIN_RIGHT, PDF_MARGIN_BOTTOM, PdfContextActive
    global PDF_CREATE_BAK, PDF_BAK_DIR, PDF_DEFAULT_BAK_DIR
    global PDF_NUM_FORMAT, PDF_FORMAT_DEFAULT, PDF_FORMAT_LIST

    PDF_START := GetPdfNum("StartNum", PDF_START)
    PDF_FONT_SIZE := GetPdfNum("FontSize", PDF_FONT_SIZE)
    PDF_MARGIN_RIGHT := GetPdfNum("MarginRight", PDF_MARGIN_RIGHT)
    PDF_MARGIN_BOTTOM := GetPdfNum("MarginBottom", PDF_MARGIN_BOTTOM)
    PdfContextActive := GetPdfNum("ContextActive", 0)

    ; 页码格式：只接受列表里的文案；旧版本没有这个键、或用户手工改坏时回落默认值
    PDF_NUM_FORMAT := GetPdfStr("PageNumFormat", PDF_FORMAT_DEFAULT)
    formatOk := false
    for item in PDF_FORMAT_LIST {
        if (item = PDF_NUM_FORMAT) {
            formatOk := true
            break
        }
    }
    if !formatOk
        PDF_NUM_FORMAT := PDF_FORMAT_DEFAULT

    ; 备份默认开启；只有明确存过 0 才关闭（旧版本没有这个键，读不到就走默认 1）
    PDF_CREATE_BAK := GetPdfNum("CreateBAK", 1) ? 1 : 0

    ; 备份目录：空串或非法时回落到默认目录，避免把文件写到奇怪的位置
    PDF_BAK_DIR := GetPdfStr("BakDir", PDF_DEFAULT_BAK_DIR)
    if (Trim(PDF_BAK_DIR) = "")
        PDF_BAK_DIR := PDF_DEFAULT_BAK_DIR

    ; v1.7：参数规范化兜底（静默/右键模式不经过 GUI 校验，这里在加载时统一收敛）
    NormalizePdfParams()
}


;---------------------------------------------------------------------
; v1.7：规范化页面参数，保证静默/右键模式拿到的数字不会越界。
; GUI 路径已有 1~999999 / 5~72 / 0~20 的校验，这里在加载注册表时也做一次兜底，
; 防止脏值绕过 GUI 校验进入 ProcessOne。
;---------------------------------------------------------------------
NormalizePdfParams() {
    global PDF_START, PDF_FONT_SIZE, PDF_MARGIN_RIGHT, PDF_MARGIN_BOTTOM

    PDF_START := Integer(Max(1, Min(999999, Integer(PDF_START))))
    PDF_FONT_SIZE := Max(5, Min(72, Float(PDF_FONT_SIZE)))
    PDF_MARGIN_RIGHT := Max(0, Min(20, Float(PDF_MARGIN_RIGHT)))
    PDF_MARGIN_BOTTOM := Max(0, Min(20, Float(PDF_MARGIN_BOTTOM)))
}

; 读字符串设置：缺失/为空时用默认值
GetPdfStr(key, defaultValue) {
    global PDF_REG_PATH

    try {
        v := RegRead(PDF_REG_PATH, key)
        if (v != "")
            return v
    }
    return defaultValue
}

; 读数字设置：注册表里缺失、为空或不是数字时用默认值，避免把脏值传给 Python
GetPdfNum(key, defaultValue) {
    global PDF_REG_PATH

    try {
        v := RegRead(PDF_REG_PATH, key)
        if (v != "" && IsNumber(v))
            return v + 0        ; +0 让 "10.5" 变成数字 10.5，而不是字符串
    }
    return defaultValue
}

SetPdfVal(key, value) {
    global PDF_REG_PATH

    try RegWrite(String(value), "REG_SZ", PDF_REG_PATH, key)
}


;---------------------------------------------------------------------
; v1.9：页码参数即时保存——控件 Change 事件调用本函数把当前值写回注册表，
; 供下次启动 LoadPdfSettings() 从注册表回填到同名控件。
; 与 RunGuiProcessing 的“校验通过后写回”以及 LoadPdfSettings 的读回语义一致：
;   StartNum / FontSize / MarginRight / MarginBottom 按各自边界归一后写；
;   PageNumFormat 走 SetPdfVal 写字符串（不合法文本直接不写）。
;---------------------------------------------------------------------
SavePdfParamNow(key, value) {
    global PDF_START, PDF_FONT_SIZE, PDF_MARGIN_RIGHT, PDF_MARGIN_BOTTOM, PDF_NUM_FORMAT

    if (key = "PageNumFormat") {
        s := String(value)
        if (s = "")
            return
        PDF_NUM_FORMAT := s
        SetPdfVal(key, s)
        return
    }
    if (!IsNumber(String(value)))
        return
    v := Float(String(value))
    if (key = "StartNum") {
        v := Integer(Max(1, Min(999999, Integer(v))))
        PDF_START := v
    }
    else if (key = "FontSize") {
        v := Max(5, Min(72, v))
        PDF_FONT_SIZE := v
    }
    else if (key = "MarginRight") {
        v := Max(0, Min(20, v))
        PDF_MARGIN_RIGHT := v
    }
    else if (key = "MarginBottom") {
        v := Max(0, Min(20, v))
        PDF_MARGIN_BOTTOM := v
    }
    SetPdfVal(key, v)
}


;=====================================================================
; [右下角浮动提示] 所有结果/错误都走这里，不弹模态框、不抢焦点
; 与公文排版的 ShowFloatingToast 同风格：深色底、白色字、右下角、自动消失。
; 高度按文字实际高度算，失败原因多行时不会被截断。
;=====================================================================
ShowPdfToast(title, msg := "", isSuccess := true, duration := 2600) {
    global PdfToastGui

    ; 同时只允许一个提示框：新提示顶掉旧提示，避免角落里叠一堆
    HidePdfToast()

    g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000", "PDF 加页码提示")
    g.BackColor := isSuccess ? "1E392A" : "5C2020"
    g.MarginX := 20
    g.MarginY := 14

    g.SetFont("s11 bold cFFFFFF", "Microsoft YaHei")
    g.Add("Text", "x20 y14 w24", isSuccess ? "✔" : "✖")
    g.Add("Text", "x+2 yp w300", title)

    msgCtrl := ""
    if (msg != "") {
        g.SetFont("s9 norm cDCDCDC", "Microsoft YaHei")
        msgCtrl := g.Add("Text", "x20 y+8 w320", msg)
    }

    ; 先让控件排版，再按消息实际高度决定窗口高度
    if IsObject(msgCtrl) {
        msgCtrl.GetPos(&mx, &my, &mw, &mh)
        toastH := my + mh + 14
    } else {
        toastH := 62
    }
    toastW := 360

    ; NoActivate：提示框不抢焦点，用户还能继续操作资源管理器/主窗口
    g.Show("w" toastW " h" toastH " NoActivate x" (A_ScreenWidth - toastW - 25)
        . " y" (A_ScreenHeight - toastH - 55))
    try WinSetTransparent(235, g.Hwnd)
    PdfToastGui := g

    if (duration > 0)
        SetTimer(HidePdfToast, -duration)
}

HidePdfToast() {
    global PdfToastGui

    if (IsObject(PdfToastGui)) {
        try PdfToastGui.Destroy()
    }
    PdfToastGui := ""
}

; 静默模式专用：没有主窗口时脚本会立刻退出，提示框会跟着消失。
; 用定时器（而不是 Sleep）让消息循环继续跑，提示框才能正常绘制，
; 到点后再按退出码结束进程。
ShowPdfToastAndExit(title, msg, isSuccess, duration, exitCode) {
    ShowPdfToast(title, msg, isSuccess, duration)
    SetTimer(() => ExitApp(exitCode), -(duration + 250))
}


;---------------------------------------------------------------------
; 统一的提示出口（静默模式）：quiet 时不提示，直接以退出码 1 结束
; （命令行/计划任务调用时不能被对话框卡住）
;---------------------------------------------------------------------
Notify(text, title, icon, quiet) {
    if quiet
        ExitApp 1

    ; 图标参数是 AHK MsgBox 的老写法，这里只用来判断是不是错误提示
    isError := InStr(icon, "!")
    ShowPdfToastAndExit(title, text, !isError, isError ? 4200 : 2600, 1)
}


;---------------------------------------------------------------------
; 把单个文件加入静默模式的待处理列表（过滤非 PDF 和 Office 锁文件）
; v1.3 起是原地修改，不会再生成 _加页码.pdf，所以不再按后缀跳过文件。
;---------------------------------------------------------------------
AddPdfFile(path, &targets, seen, &skipped) {
    SplitFilePath(path, &dir, &name, &ext)
    if (StrLower(ext) != "pdf")
        return                                   ; 文件夹里混着其他文件很正常，静默跳过
    if (SubStr(name, 1, 2) = "~$")
        return                                   ; Word/WPS 打开文档时生成的锁文件

    key := StrLower(path)
    if seen.Has(key) {
        ; v1.12：重复文件也要计数，否则静默模式的汇总里“跳过 N 个”永远是 0
        skipped += 1
        return
    }
    seen[key] := true
    targets.Push(path)
}


;---------------------------------------------------------------------
; 处理一个文件：备份 → 后端写临时文件 → 覆盖原文件
; 返回 { file, ok, out, msg, bak }，供汇总和日志使用
;
; runnerExe / runnerScript 有两种组合（命令行协议完全相同）：
;   编译版：runnerExe = 内嵌后端 pdf_page_number.exe，runnerScript = ""（不拼脚本参数）
;   源码态：runnerExe = python.exe，runnerScript = 同目录 pdf_page_number.py
;
; 为什么先写临时文件再覆盖：
;   pypdf 无法边读边写同一个文件；而且中途失败/断电时原文件必须保持完好。
;   临时文件放在原文件同目录（同一卷），最后一步 FileMove 才是真正的替换。
;---------------------------------------------------------------------
ProcessOne(inPath, idx, runnerExe, runnerScript) {
    global PDF_START, PDF_FONT_SIZE, PDF_MARGIN_RIGHT, PDF_MARGIN_BOTTOM
    global PDF_NUM_FORMAT
    global PDF_CREATE_BAK

    ; ---------- ① 可写性检查 ----------
    ; 被 Acrobat/WPS/预览窗格占用时，备份或覆盖都会失败；提前拦下，报错更清楚
    ; v1.7：先单独判文件是否存在，再判可写；FileOpen 的 "rw" 在目标不存在时会“新建”，
    ; 不防存在判断会把删掉/改名的文件“救活”成 0 字节文件，并顺带生成空备份。
    if !FileExist(inPath) {
        return FailPdfResult(inPath, "文件不存在：" inPath)
    }
    if !IsPdfWritable(inPath)
        return FailPdfResult(inPath, "文件被占用或没有写入权限，请先关闭 PDF 阅读器")

    ; ---------- ② 备份 ----------
    ; 这是原地修改唯一的回退手段：备份没成功就绝对不动原文件
    bakPath := ""
    if (PDF_CREATE_BAK) {
        b := BackupPdfFile(inPath)
        if (!b.ok)
            return FailPdfResult(inPath, b.msg)
        bakPath := b.path
    }

    ; ---------- ③ 调后端写临时文件 ----------
    tmpPath := inPath ".pgnum.tmp"
    try FileDelete tmpPath

    ; 结果文件用唯一名字，多实例/连续处理不会互相覆盖
    statusFile := A_Temp "\pdf_pagenum_" A_TickCount "_" idx ".txt"
    try FileDelete statusFile

    ; 参数全部显式转成字符串：AHK 的浮点转字符串格式不直观，Format 更可控
    ; 后端 exe 自带入口脚本，只有源码态才需要把 .py 作为第一个参数；
    ; 拼空字符串会产生一对空引号 ""，Python 会把它当成“打开空文件名”直接报错。
    ; 页码格式在界面上是中文文案，后端只认 plain/dash/total + left/center/right
    fmt := PdfFormatArgs(PDF_NUM_FORMAT)
    cmd := Quote(runnerExe)
    if (runnerScript != "")
        cmd .= " " Quote(runnerScript)
    cmd .= " --out " Quote(tmpPath)
        . " --start " Format("{:d}", PDF_START)
        . " --font-size " Format("{:.2f}", PDF_FONT_SIZE)
        . " --margin-right " Format("{:.2f}", PDF_MARGIN_RIGHT)
        . " --margin-bottom " Format("{:.2f}", PDF_MARGIN_BOTTOM)
        . " --format " fmt.style
        . " --align " fmt.align
        . " --status " Quote(statusFile)
        . " " Quote(inPath)

    exitCode := -1
    try {
        ; 注意：AHK v2 的 RunWait 用“返回值”给退出码，第 4 个参数是 PID 输出变量，
        ; 写成 RunWait(cmd,, "Hide", &code) 拿到的是 PID，会误判成败。
        ; "Hide" 避免每次处理都闪黑框；RunWait 保证一个文件处理完再处理下一个
        exitCode := RunWait(cmd, , "Hide")
    } catch as err {
        try FileDelete tmpPath
        return FailPdfResult(inPath, "无法启动后端：" err.Message)
    }

    msg := ""
    if FileExist(statusFile) {
        try msg := Trim(FileRead(statusFile, "UTF-8"), " `t`r`n")
        try FileDelete statusFile
    }

    if (exitCode != 0) {
        ; 后端失败：删掉半成品，原文件保持原样（备份也留着，不影响）
        try FileDelete tmpPath
        ; 后端在写状态文件前就崩掉时（例如解压不完整、被杀软拦下），只能报退出码
        if (msg = "")
            msg := "后端异常退出，退出码 " exitCode
        return FailPdfResult(inPath, msg)
    }

    ; 退出码 0 但没生成文件：同样视为失败，绝不动原文件
    if !FileExist(tmpPath) {
        return FailPdfResult(inPath, "后端没有生成输出文件")
    }

    ; ---------- ④ 覆盖原文件 ----------
    if !ReplaceOriginalPdf(tmpPath, inPath) {
        try FileDelete tmpPath
        return FailPdfResult(inPath, "无法覆盖原文件（可能被占用、只读或云同步盘锁定）")
    }

    result := { file: inPath, ok: true, out: inPath, msg: msg, bak: bakPath }
    AppendPdfLog("成功", inPath, PdfLogDetail(msg, bakPath))
    return result
}


; 统一构造失败结果，并立刻写一行日志（不等到批次结束，崩溃也不会丢记录）
FailPdfResult(inPath, reason) {
    AppendPdfLog("失败", inPath, reason)
    return { file: inPath, ok: false, out: inPath, msg: reason, bak: "" }
}

; 日志明细：后端提示 + 备份路径，方便事后找回原件
PdfLogDetail(msg, bakPath) {
    detail := msg
    if (bakPath != "")
        detail .= (detail = "" ? "" : "；") "备份：" bakPath
    return detail
}


;---------------------------------------------------------------------
; 可写性检查：用 rw 方式试打开再立刻关闭。
; 只读文件、被阅读器占用、没有权限都会在这里失败，比覆盖到一半再报错安全。
; v1.7：再前置 FileExist 判一次，防止 FileOpen("rw") 在文件不存在时“新建”出 0 字节文件。
;---------------------------------------------------------------------
IsPdfWritable(path) {
    if !FileExist(path)
        return false
    f := ""
    try {
        f := FileOpen(path, "rw")
        f.Close()
        return true
    } catch {
        try {
            if IsObject(f)
                f.Close()
        }
        return false
    }
}


;---------------------------------------------------------------------
; 备份原文件到备份目录：原名_yyyyMMdd_HHmmss.pdf
; 同一秒内处理同名文件时追加 (2)(3)…，绝不覆盖已有备份。
; 返回 { ok, path, msg }；失败时调用方必须放弃处理该文件。
;---------------------------------------------------------------------
BackupPdfFile(inPath) {
    global PDF_BAK_DIR, PDF_DEFAULT_BAK_DIR

    dir := Trim(PDF_BAK_DIR, " `t`"")
    if (dir = "")
        dir := PDF_DEFAULT_BAK_DIR

    ; 目录可能还没建（第一次使用或用户手工删过）
    if !DirExist(dir) {
        try DirCreate dir
        catch as err
            return { ok: false, path: "", msg: "无法创建备份目录：" dir "`n" err.Message }
    }
    if !DirExist(dir)
        return { ok: false, path: "", msg: "备份目录不存在且无法创建：" dir }

    SplitFilePath(inPath, &srcDir, &name, &ext)
    stamp := FormatTime(A_Now, "yyyyMMdd_HHmmss")
    bak := dir "\" name "_" stamp "." ext

    ; 同一秒处理同一个文件（例如重复拖入）时，给已有备份让路
    n := 2
    while FileExist(bak) {
        bak := dir "\" name "_" stamp "(" n ")." ext
        n += 1
    }

    try FileCopy inPath, bak
    catch as err
        return { ok: false, path: "", msg: "备份失败（原文件未改动）：" err.Message }

    return { ok: true, path: bak, msg: "" }
}


;---------------------------------------------------------------------
; 用临时文件覆盖原文件。
; 先 FileMove（同一卷上是原子改名，最快也最安全）；个别情况下
; MoveFileEx 被索引器/杀软短暂挡住，再退回“复制覆盖 + 删临时文件”。
;---------------------------------------------------------------------
ReplaceOriginalPdf(tmpPath, targetPath) {
    ; v1.7：FileMove（同一卷上是原子改名，最快也最安全）失败先重试几次（索引器/杀软短暂占用），
    ; 全部失败再退回“复制覆盖 + 删临时文件”。重试只盯在原子改名这一步，避免把原文件拖进危险兜底。
    replaced := false
    attempt := 0
    while (attempt < 5) {
        attempt += 1
        try {
            FileMove tmpPath, targetPath, 1      ; 1 = 允许覆盖
            replaced := true
        } catch {
            Sleep 60
        }
        if (replaced)
            break
    }
    if (replaced)
        return true

    try {
        FileCopy tmpPath, targetPath, 1
        try FileDelete tmpPath
        return true
    }

    return false
}


;---------------------------------------------------------------------
; 执行日志：程序目录 PDF加页码执行日志.txt，追加写入、不轮转
; 每行：时间<Tab>结果<Tab>路径<Tab>说明（换行/Tab 会被压成空格）
;---------------------------------------------------------------------
AppendPdfLog(result, filePath, detail := "") {
    global PDF_LOG_NAME

    logPath := A_ScriptDir "\" PDF_LOG_NAME
    line := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "`t" result "`t" filePath
    if (detail != "")
        line .= "`t" SanitizePdfLog(detail)
    try FileAppend(line "`r`n", logPath, "UTF-8")
}

; 保证一条记录严格只占一行
SanitizePdfLog(text) {
    s := String(text)
    s := StrReplace(s, "`r", " ")
    s := StrReplace(s, "`n", " ")
    s := StrReplace(s, "`t", " ")
    return Trim(s)
}


;---------------------------------------------------------------------
; 结果汇总：右下角提示框，不弹模态框
; 成功几个 / 失败几个；失败原因最多列 3 条，完整记录在日志里。
; 窗口模式调用后立刻返回（界面继续可用），静默模式由调用方等待后退出。
;---------------------------------------------------------------------
ShowPdfResult(resultList, skipped := 0) {
    info := BuildPdfResultInfo(resultList, skipped)
    ShowPdfToast(info.title, info.msg, info.ok, info.ok ? 3000 : 6000)
}

; 静默模式：显示同样的提示框，停留几秒后按退出码结束进程
ShowPdfResultAndExit(resultList, skipped, exitCode) {
    info := BuildPdfResultInfo(resultList, skipped)
    ShowPdfToastAndExit(info.title, info.msg, info.ok, info.ok ? 3000 : 6000, exitCode)
}

; 生成提示框文字：{ title, msg, ok }
; 单独拆出来，保证窗口模式和静默模式的措辞完全一致
BuildPdfResultInfo(resultList, skipped := 0) {
    okCount := 0
    failList := []
    for r in resultList {
        if r.ok
            okCount += 1
        else
            failList.Push(r)
    }

    if (resultList.Length = 0)
        return { title: "没有需要处理的文件", msg: "", ok: true }

    title := "处理完成：成功 " okCount " 个"
    if (failList.Length > 0)
        title .= "，失败 " failList.Length " 个"

    lines := []
    if (okCount > 0)
        lines.Push("页码已写回原文件" PdfBakSuffix())
    if (skipped > 0)
        lines.Push("另有 " skipped " 个文件被跳过")

    ; 只列前 3 条失败原因，避免提示框长到看不完；其余在日志里查
    Loop Min(failList.Length, 3) {
        r := failList[A_Index]
        SplitPath r.file, &fname
        lines.Push(fname "：" (r.msg = "" ? "未知原因" : r.msg))
    }
    if (failList.Length > 3)
        lines.Push("……另有 " (failList.Length - 3) " 个失败，详见日志")

    msg := ""
    for i, line in lines
        msg .= (i = 1 ? "" : "`n") line

    return { title: title, msg: msg, ok: failList.Length = 0 }
}

; 备份是否开启决定提示文案
PdfBakSuffix() {
    global PDF_CREATE_BAK

    return PDF_CREATE_BAK ? "（已备份）" : "（未备份）"
}


;---------------------------------------------------------------------
; 找 Python：先找常见安装位置，再退回 PATH 里的 python.exe
;---------------------------------------------------------------------
FindPython() {
    ; A_LocalAppData 在 AHK v2 里不存在（v1 才有），要用 EnvGet 取
    localAppData := EnvGet("LOCALAPPDATA")
    candidates := [
        "C:\Python\python.exe",
        A_ProgramFiles "\Python313\python.exe",
        A_ProgramFiles "\Python312\python.exe",
        A_ProgramFiles "\Python311\python.exe",
        localAppData "\Programs\Python\Python313\python.exe",
        localAppData "\Programs\Python\Python312\python.exe",
        localAppData "\Programs\Python\Python311\python.exe"
    ]
    for p in candidates {
        if (p != "" && FileExist(p))
            return p
    }

    ; 退回 PATH：where 会输出全部匹配项，取第一个真实存在的
    tmp := A_Temp "\pdf_pagenum_where.txt"
    try FileDelete tmp
    try RunWait A_ComSpec ' /c where python.exe > "' tmp '" 2>nul', , "Hide"
    if FileExist(tmp) {
        content := ""
        try content := FileRead(tmp, "UTF-8")
        try FileDelete tmp
        for line in StrSplit(Trim(content, " `t`r`n"), "`n") {
            line := Trim(line, " `t`r")
            if (line != "" && FileExist(line))
                return line
        }
    }
    return ""
}


;---------------------------------------------------------------------
; 依赖自检：Python 里能不能 import pypdf 和 reportlab
; 只在源码态用；编译版根本不需要本机 Python。
;---------------------------------------------------------------------
HasPdfLibs(pythonPath) {
    ; 同上：退出码是 RunWait 的返回值
    exitCode := 1
    try exitCode := RunWait(Quote(pythonPath) ' -c "import pypdf, reportlab"', , "Hide")
    return exitCode = 0
}


;=====================================================================
; [后端解析] 决定这次用哪个程序去加页码
;---------------------------------------------------------------------
; 返回 { exe, script, err }：
;   exe    = 要启动的程序（内嵌后端 exe 或 python.exe）；空字符串表示不可用
;   script = 源码态要传给 python.exe 的 .py 路径；编译版为空字符串
;   err    = 仅在 exe 为空时有内容，调用方直接拿去弹窗/写日志
;=====================================================================
ResolveRunner() {
    ; 编译版：只认内嵌后端。这里刻意不再回退去找本机 Python——
    ; EXE 里已经带着后端，找不到只说明解压/编译出了问题，
    ; 这时提示“请安装 Python”只会让用户一头雾水。
    if (A_IsCompiled) {
        info := EnsureBackend()
        if (info.exe != "")
            return { exe: info.exe, script: "", err: "" }
        return { exe: "", script: "", err: info.err }
    }

    ; 源码态（直接用 AutoHotkey 跑 .ahk）：仍走本机 Python，
    ; 这样改 pdf_page_number.py 后不用重新打包就能调试。
    py := FindPython()
    if (py = "")
        return { exe: "", script: "", err: "没有找到 Python。`n`n"
            . "源码态调试需要 Python 3 以及 pypdf、reportlab。`n"
            . "请先安装 Python（建议 C:\Python），再执行：`n`n"
            . "python -m pip install pypdf reportlab" }

    scriptPath := A_ScriptDir "\pdf_page_number.py"
    if !FileExist(scriptPath)
        return { exe: "", script: "", err: "找不到配套脚本：`n" scriptPath
            . "`n`n源码态请把 pdf_page_number.py 和本脚本放在同一个文件夹里。" }

    ; 依赖自检只做一次，避免每个文件都弹一次同样的错
    if !HasPdfLibs(py)
        return { exe: "", script: "", err: "Python 里缺少 pypdf 或 reportlab。`n`n"
            . "请在命令行执行一次：`n`n" . Quote(py) " -m pip install pypdf reportlab" }

    return { exe: py, script: scriptPath, err: "" }
}


;---------------------------------------------------------------------
; 确保内嵌后端已解压到本地，返回 { exe, err }
;
; 目录布局（版本号进路径，升级后天然隔离，不会新旧 DLL 混用）：
;   %LOCALAPPDATA%\PdfPageNumberTool\
;       backend-1.2\
;           pdf_page_number.exe
;           _internal\...          ← PyInstaller onedir 运行时依赖
;           backend.ok             ← 版本标记，最后写入
;
; 为什么用 backend.ok 判断而不是“exe 存在就行”：解压中途被断电/杀软打断时，
; exe 可能已经落地但 _internal 还缺文件，直接运行会以 0xC0000135（找不到 DLL）
; 之类的错误崩掉，而且每次启动都崩、用户看不出原因。
;---------------------------------------------------------------------
EnsureBackend() {
    global PDF_TOOL_VERSION

    exeName := "pdf_page_number.exe"
    baseDir := EnvGet("LOCALAPPDATA") "\PdfPageNumberTool"
    finalDir := baseDir "\backend-" PDF_TOOL_VERSION
    finalExe := finalDir "\" exeName
    marker := finalDir "\backend.ok"

    ; ---------- 快路径：已经展开过 ----------
    ; 标记文件里的版本号必须与当前版本一致（升级后即使同名也重新解压，
    ; 免得旧 _internal 里的 Python 运行时和新脚本对不上）
    if (FileExist(finalExe) && FileExist(marker)) {
        sameVersion := false
        try sameVersion := (Trim(FileRead(marker, "UTF-8"), " `t`r`n") = PDF_TOOL_VERSION)
        if sameVersion
            return { exe: finalExe, err: "" }
    }

    ; ---------- 需要展开 ----------
    try DirCreate baseDir

    ; FileInstall 是编译期指令：源路径 build\backend.zip 在编译时被塞进 EXE，
    ; 运行时才真正复制出来。目标名带 TickCount，多开实例同时首次运行不会抢同一文件。
    zipPath := A_Temp "\pdf_pagenum_backend_" A_TickCount ".zip"
    try FileDelete zipPath
    try {
        FileInstall("build\backend.zip", zipPath, 1)
    } catch as err {
        return { exe: "", err: "释放内嵌后端失败：" err.Message
            . "`n`n请重新获取完整的 PDF 加页码工具.exe。" }
    }
    if !FileExist(zipPath)
        return { exe: "", err: "EXE 里没有内嵌后端。`n`n"
            . "这通常是编译时漏了打包步骤（build\backend.zip 不存在），请重新编译或重新获取 EXE。" }

    ; 先解到临时目录，成功后再整体改名，这样 finalDir 永远不会停在“解压一半”的状态
    tmpDir := finalDir ".tmp" A_TickCount
    try DirCreate tmpDir
    if !ExtractZip(zipPath, tmpDir) {
        try FileDelete zipPath
        try DirDelete tmpDir, true
        return { exe: "", err: "解压内嵌后端失败。`n`n常见原因：杀毒软件拦截、"
            . "临时目录被安全软件清理、Windows 版本过旧（需要 Win10 1803+ 自带的 tar.exe）。" }
    }
    try FileDelete zipPath

    ; 标记文件最后写：内容是版本号，写成功才代表“这个目录可用”
    try FileAppend PDF_TOOL_VERSION, tmpDir "\backend.ok", "UTF-8"

    ; 目标目录里可能残留上次失败留下的半成品（没有 backend.ok），先清掉再改名
    if FileExist(finalDir)
        try DirDelete finalDir, true

    try DirMove tmpDir, finalDir

    if !FileExist(finalExe) {
        ; 可能是另一个实例抢先完成了同一版本的解压，也可能是改名失败
        try DirDelete tmpDir, true
        return { exe: "", err: "内嵌后端展开失败：找不到 " exeName "。`n`n"
            . "请手动删除下面的目录后重试：`n" finalDir }
    }

    ; ---------- 清理旧版本目录 ----------
    ; 升级后旧版本通常占 30 MB 以上；删不掉（例如旧实例还在用）就跳过，
    ; 不影响本次运行。带 .tmp 的目录是解压残留，一并清掉。
    Loop Files, baseDir "\backend-*", "D" {
        if (A_LoopFilePath = finalDir)
            continue
        try DirDelete A_LoopFilePath, true
    }

    return { exe: finalExe, err: "" }
}


;---------------------------------------------------------------------
; 用系统自带的 tar.exe 解 zip（Windows 10 1803 起内置 bsdtar，支持 zip）
; 返回 true / false，不抛异常，调用方统一处理失败提示
;
; 为什么不用 PowerShell 的 Expand-Archive：启动 PowerShell 要 0.5–1 秒，
; 部分机器还有执行策略/杀软限制；tar.exe 是原生程序，快且稳。
;---------------------------------------------------------------------
ExtractZip(zipPath, destDir) {
    tarExe := A_WinDir "\System32\tar.exe"
    if !FileExist(tarExe)
        return false

    code := 1
    try code := RunWait(Quote(tarExe) " -xf " Quote(zipPath) " -C " Quote(destDir), , "Hide")
    return code = 0
}


;---------------------------------------------------------------------
; 小工具
;---------------------------------------------------------------------
; 拆路径：目录 / 主名 / 扩展名。
; 特意不直接用 SplitPath 的第 4、5 个输出——实测这两个参数的顺序是
; “扩展名、主名”，和直觉相反，极易写错（写反了会把所有文件当非 PDF 跳过）。
; 这里只取前三个输出（完整文件名、目录），主名和扩展名自己按最后一个点拆。
SplitFilePath(path, &dir, &name, &ext) {
    SplitPath path, &fileName, &dir
    dot := InStr(fileName, ".", , -1)
    if (dot > 1) {
        name := SubStr(fileName, 1, dot - 1)
        ext := SubStr(fileName, dot + 1)
    } else {
        name := fileName
        ext := ""
    }
}

; 判断指定目录下是否存在子文件夹（用于拖入时询问是否递归）
HasSubFolders(dirPath) {
    Loop Files, dirPath "\*", "D"
        return true
    return false
}

; 给路径加双引号（路径里有空格时必需）
Quote(s) {
    return '"' s '"'
}

; 处理期间显示“忙”鼠标，避免用户以为没反应
SetCursorBusy(busy) {
    static IDC_WAIT := 32514, IDC_ARROW := 32512
    h := DllCall("LoadCursor", "Ptr", 0, "Ptr", busy ? IDC_WAIT : IDC_ARROW, "Ptr")
    DllCall("SetCursor", "Ptr", h)
}
