# 党政机关公文自动排版引擎（AutoHotkey v2）

基于 GB/T 9704-2012 的 `.doc`/`.docx` 排版工具。仓库里是**两个互不依赖的工具**；没有测试/lint/CI，也没有一键构建脚本（PDF 后端打包/编译是手工命令），Word/WPS 操作全部走 COM。

1. 排版引擎：单个按版本号命名的 AHK 脚本，当前 `6.8.33.ahk`。
2. PDF 加页码：AHK 外壳 + Python 后端。当前版本 `v1.18`（在 v1.17 三态基础上给两个下拉框加分组分隔线）；`v1.15` 与 `v1.16` 是并行回退分支（`v1.15` 的 `.exe` 已不在工作区），`v1.17` 保留作上一版回退。
3. `agents-md-for-repo.json` 是会话导出的 JSON（2.7 MB），不属于项目，别提交。

## 当前状态（2026-11-04，动手前先核对）

- 排版引擎：`6.8.29.ahk` 是现行版本，`README.md`「当前版本」与之一致。6.8.20 加了**页眉设置界面与方案字段**（三态 + 七个参数随方案保存 + 首页摘要）；6.8.21 新增 `SetDocumentHeaders()`，排版时**真正写 Word 页眉**（三态、逐节、只碰主页眉/已启用的首页与偶数页页眉），并把执行日志文件名由 `公文排版执行日志.txt` 改为 `公文.log`；6.8.22 页眉增加**下划线**，并给首页任务列表补齐右键菜单（打开文件/打开所在目录/从备份恢复/移除所选/清空列表，新增 `GuiBatchRunning` 拦截批量期间的操作）；6.8.23 把页眉下划线由字符级 `Font.Underline` 改为**段落的下边框**（`ApplyHeaderBottomBorder()`，`Range.Borders.Item(-3)` 的 `LineStyle=1`/`0`、`LineWidth=6`、`Color=0`），线横跨整个段落左右撑满，同时把字符下划线显式归零防双线；下划线复选框右移到 `x580` 与「加粗」拉开 38 像素；任务列表「结果」列不再追加「（有提示）」；6.8.24 把界面选项改名为**“页眉横线”**（方案字段仍叫 `HeaderUnderline`，复选框加宽到 `x580 w88`），实现不变，仍画段落下边框；6.8.26 页眉文字改为**默认不记住**（新增 `HeaderKeep` 字段与“记住文字”复选框，默认不勾选；启动时把未勾选方案的 `HeaderText` 写成空串，只清字段本身）；6.8.27 将“记住文字”改为红色（cFF0000），并沿用用户手工微调的布局（输入框 w250、复选框 x570 y605 w80）；6.8.28 实现**正文/标题段落样式化应用**（`TrySetParaStyle()` + 正文/标题/副标题/标题1-4 映射，混排段保持正文样式，设置样式后重施直接格式，全部 try 兼容 WPS）；6.8.29 **先调整内置样式格式再应用到段落**（`AdjustBuiltinStyleSpacing()` 把 Normal/Title/Subtitle/Heading1~4 样式的段前段后/行距调为国标，避免段落继承样式自带的非国标间距）。
- PDF 工具：`PDF 加页码工具 v1.18.ahk` + `v1.18.exe` 已完成待验收；`程序整体说明.md` §13 与 `PDF加页码项目整体情况概览.md` 都同步到 v1.18。
- **v1.15 / v1.16 是并行回退分支，别互相“合并”**：v1.15 在 v1.14 上加了页码「无」+ `--no-page-numbers` + `AddHeader` 复选框；v1.16 从 **v1.14** 另起，没有页码「无」，页眉开关是下拉首项「不添加页眉」→ `--no-header`。主线是 v1.17（页码/页眉三态）→ v1.18（下拉框分组分隔线），`v1.15`/`v1.16` 原样保留作回退。
- **v1.18 已完成待验收**：`PDF 加页码工具 v1.18.ahk`（`PDF_TOOL_VERSION="1.18"`，2,596 行）+ `v1.18.exe`（20,505,088 B）；两个下拉改成 owner-draw（`+0x210` + `WM_DRAWITEM`），页码下拉 14 项（分隔线在 0 基 3/7/11）、页眉下拉 6 项（分隔线在 0 基 3），`PDF_SEP` 占位项由 `PdfSepGuard()` 跳过；只改 AHK 外壳，`pdf_page_number.py` 与 `build\backend.zip`（19,123,458 B / 33 项）未动。验证：`/validate` EXIT=0、EXE 静默 11 页副本 exit=0、GUI 16 项检查 0 失败；剩 GUI 验收 + 重勾右键菜单。细节见概览 §2.7。
- 文档里的“当前版本”可能落后于工作区。判断真实状态看 `PDF_TOOL_VERSION`、文件时间、`build\backend.zip` 时间、`%LOCALAPPDATA%\PdfPageNumberTool\backend-<版本>\backend.ok`。
- 源码态跑 `v1.18.ahk` 用的就是当前 `pdf_page_number.py`（与 v1.18 EXE 内嵌后端一致）；跑旧版 `.ahk`（v1.13–v1.17）也会用到新版后端，行为可能和它们自己的 EXE 不同，对比时以各自 EXE 为准。

## 先读这些

- `备忘录.txt`：用户的 10 条协作约定 + 交接规则（**每次开始改代码前先更新对应的《…整体情况概览.md》**，不是改完再写）。
- `6.8.30.ahk`：唯一现行排版脚本（约 5,503 行）。要改先复制成 `6.8.31.ahk`；旧版本只作对照。
- `PDF 加页码工具 v1.18.ahk`：PDF 工具现行源码（UTF-8 带 BOM + LF，约 2,600 行）；新改动在新版本号文件上继续。
- `README.md`：用户行为说明 + 逐版本日志（「当前版本」必须跟版本走）。
- `程序整体说明.md`：架构、注册表、已知坑；第 13 节是 PDF 工具，已同步 v1.18。与脚本冲突时以脚本为准。
- `PDF加页码项目整体情况概览.md`：PDF 工具详细交接文档，含每个版本的实施记录与验证数据；接手先核对 §2 和“下一步”。

## 硬性协作约定

1. 全程中文；**先只做规划，用户明确同意后才写代码**。
2. 排版引擎每次改都**升版本号、另存新文件**；同时更新文件头 `V6.8.x`、`global APP_VERSION := "6.8.x"`（6.8.19 在第 171 行、6.8.20 在第 185 行、6.8.21 在第 202 行、6.8.22 在第 210 行、6.8.23 在第 222 行、6.8.24 在第 228 行、6.8.25 在第 233 行、6.8.26 在第 238 行、6.8.27 在第 245 行、6.8.28 在第 255 行、6.8.29 在第 263 行、6.8.30 在第 268 行、6.8.31 在第 273 行、6.8.32 在第 278 行、6.8.33 在第 278 行）、文件头「修改记录」、`README.md` 的「当前版本」。
3. PDF 工具从 v1.13 起**文件名带版本号**：新版本 = 新 `.ahk` + 新 `.exe`，并升 `PDF_TOOL_VERSION`（同时是解压目录名和 `backend.ok` 的内容）。旧版本文件（v1.12 无版本号、v1.13/v1.14/v1.16/v1.17 成品、v1.15 的 `.ahk`）保留作回退，不要覆盖；v1.18 是当前版本。
4. 升 PDF 版本后**右键菜单仍指向旧 EXE**：`RepairPdfContextMenu()` 只在旧路径失效时重写。必须在新版窗口里把「在 PDF 文件右键菜单中添加…」取消再勾选一次。
5. 代码必须写详细中文注释，重点解释 COM 调用、段落索引变化和边界条件。
6. 改完必须更新 `程序整体说明.md`；PDF 工具还要更新 `PDF加页码项目整体情况概览.md`。最后**自动启动程序**给用户测试。
7. 只有确实改了代码，回复最后一行才写 `已完成`；纯讨论或只改文档不写。
8. 旧文档里的 `PDF加页码.ahk` / `PDF加页码.exe` 指的就是 PDF 工具；现名是 `PDF 加页码工具`（**中间有一个空格**），按旧名找会扑空，也不要另建同名文件。

## 运行与验证

AutoHotkey 和 Git 都不在 PATH；没有测试套件，验证靠“跑一遍 + 看日志”。

```powershell
# 排版引擎 GUI（会弹窗交给用户测试）
& 'C:\AutoHotkey\v2\AutoHotkey64.exe' '6.8.33.ahk'

# 排版引擎静默批量（文件或文件夹；文件夹只扫第一层，会真的改文档并生成备份/PDF/日志）
& 'C:\AutoHotkey\v2\AutoHotkey64.exe' '6.8.33.ahk' '某副本.docx'

# PDF 工具：语法检查 / 开窗口 / 静默处理
cmd /c "C:\AutoHotkey\v2\AutoHotkey64.exe /validate ""PDF 加页码工具 v1.18.ahk"" & echo EXIT=%errorlevel%"
& 'C:\AutoHotkey\v2\AutoHotkey64.exe' 'PDF 加页码工具 v1.18.ahk'
& 'C:\AutoHotkey\v2\AutoHotkey64.exe' 'PDF 加页码工具 v1.18.ahk' --quiet '某副本.pdf'
```

- `AutoHotkey64.exe` 是 GUI 子系统程序：PowerShell 里 `&` 不会等它结束，`$LASTEXITCODE` 也是空的。要看退出码就套 `cmd /c "... & echo EXIT=%errorlevel%"`（已实测）或 `Start-Process -Wait`。
- 验证排版引擎：把最新 `.ahk` 和一份 `.docx` **副本**放临时目录静默跑两次，确认 `公文.log` 按“一文件一行（Tab 分隔）”追加。不要拿仓库里的示例文档直接跑，静默模式会真的改文件。
- 验证 PDF 工具：用 PDF 副本跑 `--quiet`，确认退出码 0、**原文件被原地改写**（备份目录多一份 `原名_yyyyMMdd_HHmmss.pdf`）、`PDF加页码执行日志.txt` 追加一行。A4、2cm/1.2cm、默认 10.5pt 时：页码基线 y≈34.02，left/center/right 的 x≈56.69 / 297.64 / 538.58；页眉基线 y≈807.9（与页码同 x）。可用 `pypdf` 的 `extract_text(visitor_text=…)` 核对，但要把 `Tm` 和 `CTM` 相乘后再取坐标，只读 `tm` 会得到 0。

### PDF 工具重编译（顺序不能反）

`FileInstall("build\backend.zip", …)` 的源是**编译期**相对路径，先编译后打包等于 EXE 里没有后端。

① 用 `C:\Python\Scripts\pyinstaller.exe` 和 `build\spec\pdf_page_number.spec` 生成 `build\dist\pdf_page_number\` → ② 把该目录压成 `build\backend.zip`（**zip 根下必须是 `pdf_page_number.exe` + `_internal\`**，不能再套一层文件夹）→ ③ Ahk2Exe 编译。

- 只改 `.ahk`（界面、文案、注释）→ 跳过 ①②，直接重编译 EXE。
- 改了 `pdf_page_number.py` → 必须重打 zip 再编译（v1.10、v1.13、v1.15、**v1.17** 都是这种情况；v1.16、**v1.18** 只改 `.ahk`，直接重编译）。
- 编译命令（`Ahk2Exe.exe` 在 `C:\AutoHotkey\Compiler\`，把版本号换成当前版本）：
  `Ahk2Exe.exe /in 'PDF 加页码工具 v1.18.ahk' /out 'PDF 加页码工具 v1.18.exe' /base 'C:\AutoHotkey\v2\AutoHotkey64.exe' /silent verbose /icon 'pdf_page_icon.ico'`
- 体积只能粗略判断（v1.18 20,505,088 B、v1.17 20,501,504 B、v1.16 20,927,488 B、v1.15 20,928,512 B、v1.14 20,925,440 B、v1.13 21,303,296 B）；**真正的判据是首跑解压出的 `backend-<版本>\backend.ok` 内容＝版本号，且静默跑副本 exit=0**。
- `Ahk2Exe.exe` 也是 GUI 子系统程序；要判断结果用 `Start-Process -Wait` 或 `[Diagnostics.Process]::Start()` + `WaitForExit()`，再看 EXE 的体积/时间。

## 排版引擎：只有踩过才知道的点

- 流水线顺序在 `FormatWordDocument()`：占用检查 → 备份 → COM（**WPS 优先**：`Kwps.Application` → `WPS.Application` → `Word.Application`）→ 编号转文本 → 页边距 → 文档网格 → 清空段 → 全文正文格式 → 标题区/附件/落款 → 表格 → 页码 → 保存/导出 PDF。
- `ApplyDocumentGrid()` 必须在页边距之后调用：改边距会让 Word/WPS 重新推算网格并覆盖先前的设置。写 `CharsLine` 会把 `LayoutMode` 改回 1，所以“只指定行网格”分支绝不能碰它；`LinesPage := 0` 会抛 `RPC_E_SERVERFAULT` (0x80010105)。
- Word 把分页符/分节符存成段落文字里的 `Chr(12)`，还可能单独设 `PageBreakBefore`：删空段或改文字前必须走 `IsPageBreakProtected()` / `HasPageBreakChar()`，改写文字统一用 `SetParagraphText()`。别把 `TrimAllSpaces()` 当判空唯一依据（`\s` 会匹配 `Chr(12)`）。
- 插入/删除段落后索引全变，必须重新取 `doc.Paragraphs()`；核心引擎里大量 `idx` 调整就是这个原因。
- WPS 兼容是硬约束：`PageSetup` 没有 `PageNumbering`，页码重置必须先/同时用 `Footers(x).PageNumbers`；不要“清理”这些兜底 `try`。
- 页码格式靠字符串 `InStr` 匹配（“无”“格式数字”“加总页码”“居左/居右”），**改下拉文案必须同步** `SetPageNumbers()` / `WriteFooterPageBlock()`。本程序写的页码块用书签 `DocFmt_PN_*` 标识，重排时只删自己的块。
- 新增方案字段要同时加到 `GetDefaultSchemeMap()`、`GetFontKeys()`（或 `GetProfileKeys()`）和编辑器/保存逻辑，否则读不到也存不进 `Profiles`。
- 数字框的上下箭头必须用 `MakeSpin()` 创建，`UpDown` 选项里**不要写裸数字**（尤其 `-2`，那会减掉 `UDS_SETBUDDYINT`，表现为箭头能点但输入框文本不变）。`MakeSpin()` 必须紧跟对应 Edit 创建，否则 buddy 绑错；它还会把初始值写进 Edit，所以调用后要恢复精确原值（50 会覆盖边距、25.4 会被截成 25）。
- 方案编辑器坐标大量写死（左列 `x20 w360`、右列 `x400 w500`、组框底边 `y564`、页眉组 `y576 h100`、窗口 `w920 h744`）；给 `GetFontItemsDef()` 加行或给左侧加控件，都要同步下移后续组框/按钮，否则控件重叠。
- 注册表全部在 `HKCU\Software\DocFormatTool` **根键下的值**；`GetConfigVal(section, key, ...)` 的 `section` 参数已被忽略，不要再按子项去查。用户方案在 `Profiles\<方案名>`。没有 `config.ini`，也不要再引入。
- 日志 `公文.log`（v6.8.21 起；旧名 `公文排版执行日志.txt` 不再写入）：追加、不轮转，`时间<Tab>结果<Tab>路径[<Tab>说明]`，结果只有 `成功`/`成功（有提示）`/`失败`。备份默认 `A_MyDocuments\公文自动排版文件备份`。
- `#SingleInstance Off` + `%TEMP%\DocFormatTool_TaskQueue.txt` 抢锁：锁只保护“读队列并清空”，**不保护排版过程**，抢不到锁直接退出。

## PDF 工具：只有踩过才知道的点

- 原地写回顺序不能变：`FileExist()` → `IsPdfWritable()` → `BackupPdfFile()` → 后端写 `%TEMP%\pdf_pagenum_<PID>_<TickCount>_<序号>.pdf` → `FileCopy` 到原目录 `原名.pdf.pgnum.tmp` 并 `FileGetSize` 校验 → `ReplaceOriginalPdf()`（`FileMove` 重试 5 次、间隔 60 ms，全失败才 `FileCopy` + `FileDelete`）。任一步失败都删临时/中间文件，原文件一个字节不动；备份失败直接放弃该文件。
- **写回必须经 `%TEMP%` 中转（v1.14），不要改回“后端直接往原文件旁边写”**：本机 Sandboxie 的 `[DefaultBox]` 有 `OpenFilePath=w:\Eyes\Desktop\公文`（还有 `c:\python`）。命中该前缀的目录里，**冻结后的** `pdf_page_number.exe` 写出的文件会落到另一层文件视图——写成功了，但外壳 `FileExist()` 看不到，日志只报“后端没有生成输出文件”；源码态走 `C:\Python\python.exe` 反而正常，所以表现为“EXE 失败、.ahk 成功”。
- 入口分流：无参数才建窗口；带文件参数（拖到 EXE、右键菜单、命令行）走 `RunSilent`；`--quiet` 完全静默，只用退出码（0 全成功 / 1 有失败）。窗口模式拖入文件夹会问是否递归，拖到 EXE 的静默模式只扫第一层。
- 提示统一走 `ShowPdfToast()`；只有 3 处允许模态框：拖入含子文件夹时问是否递归、恢复上一版本的覆盖确认、`DirSelect` 选目录。状态栏只写事实，**不要写“结果已同时在右下角提示”这类元描述**（用户明确嫌傻）。
- 静默模式没有窗口时靠 `SetTimer` 保活再 `ExitApp`，**不能用 `Sleep`**（消息循环停住，toast 画不出来）。
- 设置存在独立键 `HKCU\Software\PdfPageNumberTool`，**不要和 `DocFormatTool` 混用**；备份/日志名是 `PDF加页码文件备份` 和 `PDF加页码执行日志.txt`（和排版引擎不同，别顺手“统一”）。
- 判后端可用看 `backend.ok`，不能只看 exe 在不在（解压中断时 exe 已落地但 `_internal` 不全）。升级会重新解压到 `backend-<版本>`，旧的 `backend-*` 只是**尝试**清理，残留属正常；`PDF_TOOL_VERSION` 一改，首跑会重新解压（慢约 2.5 秒）。
- 联动置灰统一由 `ApplyPdfUiState()`（v1.17 起同时管页码三项、字号和页眉两项；v1.16 是 `ApplyPdfHeaderState()`；v1.15 的 `ApplyPdfUiState()` 只管「无」/复选框）计算，它**必须先看批量锁 `PdfUiLocked`**，否则处理期间会被它重新“启用”一部分控件；批量结束解锁后要重新应用一次。
- v1.17 页码/页眉各三态，用哨兵文案判断：`PDF_PAGE_KEEP`（不改变页码）/ `PDF_PAGE_CLEAR`（清空页码）/ `PDF_HEADER_KEEP`（不改变页眉）/ `PDF_HEADER_CLEAR`（清空页眉）。`ProcessOne()` 据此拼 `--keep-page-numbers` / `--no-page-numbers` / `--keep-header` / `--no-header`；只有“添加”才传 `--header*` 和 `--format/--align`。页眉文字留空在后端含义是“用 PDF 文件名”，**不能拿空串表示“不加/不改变页眉”**。页眉文字允许空串，必须用 `GetPdfRawStr()` 读取；`GetPdfStr()` 会把空串替换成默认值。v1.16 的「不添加页眉」已不在新列表里，注册表旧值加载时回落到「不改变页眉」。
- v1.18 下拉框分组：`PDF_FORMAT_LIST` / `PDF_HEADER_ALIGN_LIST` 仍是纯数据（11 / 5 项），界面条目由 `PdfAddSeps()` 生成 `PDF_FORMAT_ITEMS`（14 项，分隔线在 3/6/9 之后）/ `PDF_HEADER_ITEMS`（6 项，分隔线在 3 之后）；`PDF_SEP := Chr(0x2500)` 只画 1px 灰线。两个下拉必须保留 `+0x210`（`CBS_OWNERDRAWFIXED|CBS_HASSTRINGS`）并在创建前 `OnMessage(0x002B, PdfSepDraw)`；`DRAWITEMSTRUCT` 用 64 位偏移（`hwndItem` 24、`hDC` 32、`RECT` 40、`itemData` 56）。`PdfSepGuard()` 键盘按方向跳、鼠标点击退回原值；分隔条不能塞进 `*_LIST`。v1.17 页眉列表里遗留的 `"-"` 已删。
- v1.13–v1.16 页眉与页码画在同一张 overlay、只 `merge_page` 一次；v1.17 改成独立指纹块（页码 `(1.13,1.33)`、页眉 `(1.13,1.32)`，旧合并层 `(1.13,1.31)`），只清一项时按块内文字视觉 y 拆分（最上=页眉、最下=页码），不要整块删。页眉**叠加在最上层**，原文顶部 1.2cm 内有内容就会重叠，无法自动避让。字号与页码共用。
- 旧叠加层清理：先找不可见标记 `DocFmtPN` 整块删除，v1.10-v1.12 的旧层才用“仅末尾页码形状”兜底。**不要退回按字号扫全内容流**——10.5pt 五号正文会被误删。
- v1.17「清空」除了删本工具旧层，还删文字层里可识别的原生页码/页眉：页脚 4cm 带内、整段文字形状匹配页码（数字与物理页号或 PDF 页标签一致；`a / b` 还要求 b=总页数）才删；页眉只删顶部 4cm 带内跨页重复的文字（单页标题不删）。扫描图片里印着的删不掉；宁可漏删，不要按字号扫全文。
- 往 `SetPdfControlsLocked()` 的锁定列表加控件时，`global` 声明必须一起加（v1.12 的教训：漏了 `uiPdfBakOpen`，每次点「开始」必崩）。
- 界面坐标是写死的（窗口 `w700 h800`，组框 `y75/225/340/505`，按钮行 `y730`）：改布局先按 10 号微软雅黑汉字约 13.3px 算标签宽度，给窄了会自动换行压住下一行。
- 已知坑：`PdfToastGui` 必须在脚本顶部初始化成空字符串，否则第一次 `HidePdfToast` 会卡在“变量未赋值”的错误框；窗口里的「开始添加页码」是 `Static` 控件，`BM_CLICK` 无效，`WM_DROPFILES` 也不能跨进程伪造，验证 GUI 只能在脚本副本里加临时钩子直接调 `PdfAddToGui()` + `RunGuiProcessing()`。
- v1.15/v1.16 踩过：AHK v2 的 `Gui.Button` **没有** `ToolTip()` 方法（`ctrl.ToolTip("…")` 直接抛错、窗口建一半崩掉）。悬停提示只能 `OnMessage(0x0200)` 或 `WM_MOUSEMOVE` + `ToolTip()`，并用一次性 `SetTimer` 自动收起；`WindowFromPoint` 在窗口被遮挡时会返回别人的句柄。

## 环境陷阱（本仓库特有）

- **文件工具的写入不一定会落到真实磁盘**：本机出现过 read/edit 工具能看到新内容、PowerShell 看到旧内容的情况。改完关键文件必须用 shell 复核（`Select-String` / 字节数 / 时间戳），以 shell 为准；批量改文件更稳的做法是“把新内容写成临时文件，再用 PowerShell 替换”。
- 通过 shell 传中文容易出问题：命令里的中文字面量可能被编码破坏、`Get-Content` 直接看中文会乱码。先 `[Console]::OutputEncoding=[Text.Encoding]::UTF8`；需要精确文本时用 `[IO.File]::ReadAllText/WriteAllText` 并显式指定编码。
- 保持编码与行尾：`6.8.29.ahk` 是 UTF-8 无 BOM + **CRLF**；`README.md` / `程序整体说明.md` / `AGENTS.md` 是 UTF-8 无 BOM + LF；`PDF 加页码工具 v*.ahk` 是 UTF-8 **带 BOM** + LF；`pdf_page_number.py` 是无 BOM + LF；`备忘录.txt` 和 `PDF加页码项目整体情况概览.md` 是 CRLF。改完用 `[IO.File]::ReadAllBytes()` 看前三字节（BOM 为 `239,187,191`）并统计 CRLF/LF，别把整个文件的行尾换掉（会产生全文件 diff）。
- 函数名 `Rd` 之类和 PowerShell 别名冲突（`rd` = `Remove-Item`），写脚本时别用短名。

## Git

- 远程 `origin` = `https://github.com/ksgujie/OfficialDocument.git`，分支 `main`（2026-11-04 由 `https://github.com/ksgujie/gongwen.git` 改为新地址；新仓库初始为空，首次推送已新建 `main`）。
- 工作区现状（2026-11-07）：已提交 `6.8.15`、`6.8.20`–`6.8.28`（本次将提交 `6.8.29`），以及 PDF 工具 `v1.12`–`v1.18` 的源码、`pdf_page_number.py`、`pdf_page_icon.ico` 和 `PDF加页码项目整体情况概览.md`；`6.7.5`-`6.8.14` 共 20 个旧版本仍显示为已删除（`D`）；未跟踪 `6.8.16`-`6.8.19`、`RULES.md`、`__pycache__/`、`agents-md-for-repo.json`。`.exe`、`*.zip`、`/build/` 被 `.gitignore` 忽略，**不会**出现在未跟踪列表里。这是用户有意清理，**不要 `git add -A`，不要提交那些 `D`**；需要旧版本用 `git show HEAD:<文件>`。
- 每个版本单独提交为一个 `.ahk` 文件；只 stage 明确要提交的文件。
- GitHub 直连经常失败。**不要永久配置代理**；探测 Clash 端口后只在本次命令加 `-c`：
  ```powershell
  Test-NetConnection 127.0.0.1 -Port 7890 -InformationLevel Quiet -WarningAction SilentlyContinue
  & 'W:\git\bin\git.exe' -C '<仓库目录>' -c http.proxy=http://127.0.0.1:7890 -c https.proxy=http://127.0.0.1:7890 push
  ```
- 推送是否成功以 `git log --oneline origin/main -1` 为准；git 正常输出到 stderr 也会显示红字，别只看颜色。
- `.gitignore` 已排除 `*.reg`、`*.doc(x)`、`*.pdf`、日志、`*.exe`、`*.zip`、`/build/`；`__pycache__/` **没有**被排除。

## 修改检查清单

- [ ] 排版引擎复制成新版本号文件，未直接改旧版本；PDF 工具新建带版本号文件并升 `PDF_TOOL_VERSION`（v1.12 文件保留）
- [ ] 更新文件头版本号、`APP_VERSION` / `PDF_TOOL_VERSION`、文件头「修改记录」
- [ ] 新代码有详细中文注释（COM、索引、边界）
- [ ] 更新 `README.md` 的「当前版本」（排版引擎）
- [ ] 更新 `程序整体说明.md`，PDF 工具还要更新 `PDF加页码项目整体情况概览.md`
- [ ] 改页眉时检查 `HeaderUnderline`（页眉横线，段落下边框 `ApplyHeaderBottomBorder()`）与置灰/清理列表；改列表右键时检查 `GuiBatchRunning` 与“从后往前删”
- [ ] 临时目录 + 副本静默跑一遍，检查日志/备份/原地改写
- [ ] 确认文件编码与行尾未变（BOM/CRLF）
- [ ] PDF 工具升版本后重新勾选右键菜单，确认指向新 EXE
- [ ] 启动 GUI 供用户测试
- [ ] 确实改了代码时，回复最后一行写 `已完成`
