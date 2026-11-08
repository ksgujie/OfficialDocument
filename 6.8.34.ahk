; 声明脚本所需的最低 AutoHotkey 版本为 v2.0
#Requires AutoHotkey v2.0
; 允许多开运行多个实例
#SingleInstance Off
; 关闭未初始化变量的警告提示
#Warn VarUnset, Off

; =====================================================================
; 党政机关公文自动排版引擎 (基于 GB/T 9704-2012) - V6.8.34 逐行注释版
;
; 修改记录：
; v6.6.1 修复页码格式
; v6.6.5 新增统一版本号变量、标题最大字数拦截阈值设置、关于窗口
; v6.6.7 移除"生成修改对比副本"功能，修复附件说明缩进覆盖问题与右键恢复菜单
; v6.6.8 新增 SetParagraphText 安全替换函数，根治文末成文日期段落分裂产生双日期的 Bug；
;        修正文末落款单位字体为正文字体（三号仿宋）
; v6.6.9 移除落款全角空格填充，改用动态计算右缩进字符数实现上下居中（单位长时右缩进 2 字符）
; v6.7.0 附件及正文各级标题（一至四级）序号及半角点"."后方的空格全部自动剔除；
;        附件说明采用基于当前方案正文字号（ptBody）的动态磅值悬挂缩进
; v6.7.1 1. 标题最大字数默认拦截阈值调整为 35 字（自动升级旧版默认 50 字配置）
;        2. 统一一至四级标题判定：末尾带标点且总长未超阈值的独立短段直接识别为标题
;        3. 新增 Word/WPS 自动编号转纯文本与特殊不可见空白清理，防止自动编号标题漏判
; v6.7.6 1. 行距预设改为与 Word 界面一致的"单倍行距"，并同步加固行距解析器
;        2. 补全文件头常量的中文尾注释（行距规则 0/4/5、边距预设 1/2/3 等）
;        3. 导出 config.ini 时补写 OptProtectTable，避免导出后该项回落默认值
;        4. 统一行距与磅值的空格写法
; v6.7.7 1. 页码写入改为安全模式：保留页脚原有内容，仅替换本程序生成的页码块；
;        2. 同时处理主页脚、首页页脚和偶数页页脚，页脚含其他内容时保留原内容并给出提示；
;        3. PDF 导出失败时在文件状态、悬浮通知和执行日志中明确提示，不再静默失败；
;        4. 新版本按版本号命名（6.7.7.ahk），并启用 Git 版本库管理。
; v6.7.8 1. 修复关闭“排版前生成原始备份”后，备份存放目录的“浏览”“打开”按钮被置灰的问题；
;        2. 备份目录的路径、浏览和打开按钮始终可用（仅在批量排版期间临时锁定）；
;        3. 新增任务列表双击文件用系统默认程序打开功能，方便排版前核对原文件；
;        4. “打开”备份目录时若目录不存在则自动创建，避免空目录提示。
; v6.7.9 1. 修复“浏览...”备份目录只能在该目录子树内选择的问题，改用系统文件夹选择框，
;        支持向上一级及任意磁盘导航；
;        2. “排版前生成原始备份”由界面复选框移入主菜单“设置”，避免误点关闭；
;        3. 开始排版前若备份开关处于关闭状态，弹出“是/否”提醒对话框（默认“是”），
;        选择“是”即重新开启备份并继续排版（界面批量与静默排版均生效）；
;        4. 批量排版期间同步禁用主菜单中的备份开关，防止运行中误改。
; v6.8.0 1. 停用 config.ini，全部设置只保存在注册表 HKEY_CURRENT_USER\Software\DocFormatTool；
;        2. 自定义字体方案同样只读写注册表，不再与配置文件双向同步；
;        3. “导出设置”改为导出 .reg 注册表文件，换一台电脑双击导入即可恢复全部设置；
;        4. 状态栏改为显示“设置保存在注册表”；
;        5. 导出文件的默认位置改为程序所在目录（编译后为 EXE 目录，未编译为脚本目录），
;        文件名固定为“公文排版设置.reg”，不再附加日期时间。
; v6.8.1 1. 导出设置成功后的模态对话框改为右下角浮动提示框，不再打断当前操作；
;        2. 浮动提示框按文字内容自动调整高度，较长的文件路径不会被截断。
; v6.8.2 1. 拖入文件夹时先检查是否包含子文件夹；若包含，弹出“是/否”提示框询问
;        是否扫描子文件夹中的文件（默认“是”）；
;        2. 选择“是”时递归添加子文件夹中的 Word 文件，选择“否”时只添加第一层。
; v6.8.3 1. 执行日志改为固定保存在程序所在目录（编译后为 EXE 目录，未编译为脚本目录）；
;        2. 日志固定为一个文件“公文排版执行日志.txt”，文件名不再附加日期时间，每次覆盖写入。
; v6.8.4 1. 新增“整体配置方案（单位/文种）”：边距、页码、字体方案、排版选项、标题阈值
;        可整体保存为一套方案，首页下拉一键切换；
;        2. 首页只显示当前方案信息和任务列表，具体参数统一在“编辑参数”窗口修改；
;        3. PDF、排版前备份、备份目录不随方案切换，单独放在首页“输出与备份”区域；
;        4. 标题阈值从菜单移入方案编辑窗口，随方案保存；
;        5. 备份开关从主菜单移到首页“输出与备份”区域；
;        6. 内置只读方案“标准公文格式”作为恢复入口，可“另存为”后编辑；
;        7. 升级后首次启动自动把现有注册表设置保存为“默认方案”，无需重新配置。
; v6.8.5 1. 修复点击“保存参数并运行”时报错“Too few parameters passed to function.
;        Specifically: Menu.Prototype.Enable/Disable”，导致排版无法开始的问题；
;        2. 批量排版期间改为逐项禁用“设置”菜单中的操作项，排版结束后恢复。
; v6.8.6 1. 内置只读方案更名为“党政机关公文格式”（原“标准公文格式”）；
;        2. 取消升级时自动创建“默认方案”，没有用户方案时直接使用内置方案；
;        3. 方案编辑窗口中，边距预设、四个边距输入框和字体方案若与国标不一致，
;        对应输入框显示为红色，方便一眼看出哪些参数偏离国标。
; v6.8.7 1. 改为“选项名称后显示红色星号”：编辑方案窗口中边距预设、上下左右边距、
;        字号方案的名称，以及字体方案管理窗口中各级标题的名称，与国标不一致时
;        名称后出现红色“*”，名称与输入框本身不变色；
;        2. 字体方案管理窗口逐项对照国标（字体、字号、加粗、正文行距）；
;        3. 启动时固定按最后一次使用的方案生效，右键/静默排版同样以该方案为准。
; v6.8.8 1. 字体字号不再单独保存为字体方案，改为与页边距等参数一起放在
;        同一张“编辑方案参数”窗口中，全部随整体方案一起保存和切换；
;        2. 编辑窗口新增各级标题的字体、字号、加粗和正文行距，逐项对照国标
;        显示红色星号；字体字号管理子窗口取消；
;        3. 旧方案仍按原字体方案名称自动迁移字体参数，无需重新设置。
; v6.8.9 1. 首页方案信息中的边距改为“37/35/28/26 mm”形式，不再显示上下左右；
;        2. “表格保护与字号适配”真正生效：勾选后不改动表格原有格式，
;        取消勾选时按方案统一表格字体、字号和对齐；
;        3. 编辑方案参数窗口删除底部“字体字号随当前方案一起保存…”说明文字；
;        4. 备份目录的“浏览…”“打开”按钮改为等高对齐；
;        5. 任务列表标题改为“排版任务列表 (拖拽文件、文件夹至此)”；
;        6. 状态栏不再显示“设置保存在注册表”字样；
;        7. 底部“清空列表”“导出设置”“结束后台排版进程”三个按钮统一宽度；
;        8. “排版增强与细节控制选项”的三个选项增加鼠标悬停说明。
; v6.8.10 1. 修复三个排版选项悬停时不显示说明的问题，改为程序自己控制的悬停提示；
;        2. 备份目录输入框上移 3 像素，“浏览…”“打开”两个按钮改为与其等高并垂直对齐；
;        3. 首页底部四个按钮重新分配宽度：“清空列表”“导出设置”变短，“结束后台排版进程”
;           保持不变，“保存参数并运行”加长。
; v6.8.11 1. 执行日志改为追加写入，不再每次覆盖；文件名仍固定为“公文排版执行日志.txt”，
;           不添加日期、时间；
;        2. 每个文件处理完立即写一行日志，成功、失败、成功但有提示都记录，
;           全部成功时也会写日志；
;        3. 日志行采用“时间 + 结果 + 文件路径 + 说明”四列，字段之间用 Tab 分隔，
;           说明中的换行和 Tab 会替换为空格，确保一个文件严格只占一行；
;        4. 状态栏和右下角提示会显示日志是否追加成功，日志写入失败时给出明确提醒；
;        5. 首页底部“导出设置”右侧新增“日志”按钮，点击直接用默认程序打开执行日志；
;           底部五个控件重新分配宽度，仍保持一行排列。
; v6.8.12 1. 标题区改为从“大标题”下一段开始连续向下扫描：凡末尾没有标点、
;           且不是一至四级标题的短段，全部作为副标题（楷体三号、居中），
;           标题与成文日期之间的空行不再保留；
;        2. 大标题默认为第一个非空段；若第一段只是“附件1”“附件 1”这类
;           单独成行的附件编号，则把紧随其后的段落作为大标题；
;        3. 扫描遇到表格、标点结尾、一至四级标题、附件说明或超长正文即停止，
;           避免把正文段落误当成副标题。
; v6.8.13 1. 修复排版时误删分页符/分节符的问题：只含分页符的段落不再当成空行删除，
;           改写段落文字时保留段末的分页符/分节符，“段前分页”的段落同样受保护；
;        2. 附件标记“附件1”按 GB/T 9704-2012 7.3.8 排版：3 号黑体、顶格、取消首行缩进，
;           字体字号在“编辑方案参数”窗口新增的“附件标记”一行里自由修改；
;        3. 附件标记与附件标题之间自动空一行，使附件标题落在版心第三行
;           （原来标题紧接第二行）；附件标题仍跟随方案的“主标题”字体并居中；
;        4. 新增方案选项“附件另起一页”：勾选时在“附件1”标记段设置段前分页，
;           实现附件另面编排；标记本身已带分页符时不重复加页，避免出现空白页；
;        5. 首页方案信息增加“附件另面”状态显示。
; v6.8.14 1. 编辑窗口“字体字号”表格重排：附件标记移到最下面，并在其下新增
;           “页码”一行（字体/字号/加粗可改），默认按国标 宋体 四号 不加粗；
;        2. 页码字体字号真正生效：页脚页码块、以及页脚已有页码域的字体
;           都改用方案里的 PageNum 参数，不再写死宋体四号；
;        3. 编辑窗口控件排版统一：左侧三个组框统一宽度与内边距，边距区做成
;           两列等宽网格，“标题阈值”并入排版选项组，字体表六列对齐、行距
;           收紧，并在“附件标记”前加一条细分隔线；底部按钮右端对齐。
; v6.8.15 1. 勾选“附件另起一页”时，改用分节符（下一页）将附件划分到新节，
;           断开新节页脚链接并将页码重置为从 1 开始，实现“附件页码单独从 1 编号”。
; v6.8.16 1. 新增文档网格设置：当前方案为内置只读方案“党政机关公文格式”，或方案
;           的边距预设选的是“党政机关公文格式”（MarginType = 1）时，版心固定为
;           每页 22 行、每行 28 个汉字（GB/T 9704-2012）；
;        2. 其余方案（WORD 默认边距、用户自定义边距）的每页行数与每行字符数
;           改用 Word 新建文档的默认值：每页 44 行、每行 39 个字符；
;        3. 网格在页边距设置完成之后再写入，因为改动边距会让 Word/WPS 按新版心
;           重新推算网格，先写会被覆盖；各属性逐项 try，兼容 WPS 对个别网格
;           属性报错的情况。
; v6.8.17 1. 新增方案参数“文档网格”：网格类型（不指定网格 / 只指定行网格 /
;           指定行和字符网格）、每页行数、每行字符数，和边距、字体一样随方案保存，
;           在“编辑方案参数”窗口底部新增的一组控件里修改；
;        2. 网格默认值按边距预设推算：边距预设为“党政机关公文格式”→ 指定行和字符
;           网格、22 行、28 字（GB/T 9704-2012）；其余预设 → 只指定行网格、44 行、
;           每行字数不指定（Word 新建文档的默认口径）；
;        3. 排版时网格改由方案字段驱动（ApplyDocumentGrid），仍放在页边距之后写入，
;           避免被 Word 按新版心重新推算覆盖；
;        4. 切换“边距预设”时，网格若仍是原预设的推荐值则自动跟随新预设，
;           手工改过（例如改成 30 行）则保持不动，不覆盖用户的设置；
;        5. 旧方案（6.8.16 及以前保存的）没有这三个键，读取时按第 2 条规则自动补齐，
;           升级后排版结果与 6.8.16 保持一致。
; v6.8.18 1. “编辑方案参数”窗口左侧重排：原来的“版心边距设置”与底部的“文档网格”
;           合并为一个名为“页面设置”的组框，组内分“页边距”“文档网格”两段，
;           “版心边距”改称“页边距”；页码设置、排版选项随之下移，窗口由 h650 收缩到 h590；
;        2. 左侧第三个组框改名为“排版选项”（少一个勾选项后名字更直白）；
;        3. 删除“规范发文字号六角括号”功能：不再有该开关、不再读取/保存 OptFixDocNum，
;           排版时也不再把开头发文字号的括号替换成〔〕（旧注册表值保留但不生效）。
; v6.8.19 “编辑方案参数”窗口视觉重排（只动坐标与样式，不动任何排版逻辑）：
;        1. 统一栅格：左右外边距 20、两列间距 20，左列 x20 w360、右列 x400 w500，
;           方案名称输入框右端与左列边界对齐；组框内控件统一左右各留 12；
;        2. 上下平衡：左列三个组框与右侧“字体字号”组框底边统一落在 y564，
;           窗口由 h590 加高到 h632，按钮行下移到 y578 并统一 h34；
;        3. 节奏统一：页边距行距 26→30、文档网格行距 26→30、字体行距 30→36，
;           两处细分隔线宽度撑满组框内宽；
;           左列增加“页边距 (毫米)”“文档网格”两个灰色小节标题，与右侧表头同款；
;        4. 边距输入改成真正等宽两列（列 1 x32、列 2 x200），标签/星号/输入框/单位
;           四段位置固定，右边距不再一边 12 一边 26；
;        5. 修复“微调按钮（上下箭头）点了没反应”：六个数字框的 UpDown 选项里都写了
;           "-2"，而 AHK 会把选项中的负数当成“从窗口样式里减掉该位”，-2 正好减掉
;           把数值写回输入框的 UDS_SETBUDDYINT(0x2)，于是箭头只改内部值、Edit 文本
;           不动。改为统一的 MakeSpin() 创建（不再写 -2，并按当前值初始化）；
;           同时补齐网格切换的推荐值：选“只指定行网格”时行数补 44（Word 默认口径），
;           选“指定行和字符网格”时行数/字符数补 22/28（国标口径）。
; v6.8.20 新增“页眉设置”（本版只做界面与方案保存，排版时还不会写 Word 页眉）：
;        1. “编辑方案参数”窗口底部新增通栏组框“页眉设置”（x20 y576 w880 h100），
;           四列栅格：列起点 32/246/460/674、标签宽 60，控件起点 96/310/524/738、
;           控件宽 134；两行 y600 / y634。窗口由 h632 加高到 h744，按钮行下移到 y690；
;        2. 页眉“处理方式”三态：不改变（保留原有页眉）/ 添加·替换 / 清空；
;           只有“添加·替换”时其余六个页眉控件可用，另外两种由
;           UpdateHeaderControlState() 整体置灰（含数字框的 UpDown 箭头）；
;        3. 页眉文字留空是合法值，含义是“用文档文件名（不含扩展名）作页眉”；
;           空输入框里用 EM_SETCUEBANNER 显示灰色提示，不额外占控件、不影响取值；
;        4. 页眉参数随方案保存：HeaderMode / HeaderText / HeaderAlign /
;           HeaderMarginTop / HeaderFont / HeaderSize / HeaderBold；页眉字体
;           不进右侧“字体字号”国标表（国标没有页眉字体规定，避免永久红星）；
;        5. 首页“配置方案”组加高到 h190，新增一行“页眉处理”摘要，其下
;           “输出与备份”、任务列表、底部按钮整体下移 25 像素，窗口 h600 → h625。
; v6.8.21 1. 新增 SetDocumentHeaders()：排版时真正写入 Word 页眉（在写页码之前调用），
;           按方案“处理方式”三态处理每一节：不改变（一个节都不碰）/
;           添加·替换（整段替换成方案文字）/ 清空（清空文字并删除页眉里的域）；
;        2. 只写主页眉 Headers(1)；仅当该节已经启用“首页不同”/“奇偶页不同”时，
;           才顺带处理 Headers(2)/Headers(3)，不擅自打开这两个版式开关；
;           写前先 LinkToPrevious := false，避免改到上一节；
;        3. 页眉文字留空时取文档文件名（不含扩展名）；对齐、字体、字号、加粗
;           按方案写入，中文字体同时写 Name/NameFarEast/NameAscii；段落缩进、
;           行距、段前段后归零，避免继承正文首行缩进；
;        4. “添加·替换”时 PageSetup.HeaderDistance 跟随方案“距上边”（默认 15 mm）；
;           “不改变”“清空”不再改写页眉距离（v6.8.20 及以前固定写 15 mm）；
;           页眉写入失败只并入 warn_msg，不中断整个文件的排版，
;           日志显示“成功（有提示）”；
;        5. “添加·替换”遇到原页眉已有内容（文字/图片/表格）时直接替换，
;           只在日志说明里提示一句，不弹窗、不跳过；
;        6. 执行日志文件名由“公文排版执行日志.txt”改为“公文.log”（旧文件保留、
;           不再写入）；AppendExecutionLog()、OpenLogFile() 与状态栏提示同步修改。
; v6.8.22 1. 页眉新增“下划线”选项：方案字段 HeaderUnderline（0=无、1=有），
;           界面放在“加粗”右侧（复选框自带“下划线”文字，x550 y638 w76）；
;           写页眉时通过 Font.Underline 生效（Word 里 -1/0 与 true/false 都接受，
;           这里统一用布尔值），摘要行在“添加·替换”下追加“· 下划线”；
;        2. 首页任务列表右键菜单参照 PDF 加页码工具补齐：
;           打开文件 / 打开所在目录 / 分隔线 / 从备份恢复此文件 / 分隔线 /
;           移除所选文件 / 清空列表；新增 OpenSelectedLVFolder()、
;           RemoveSelectedLVItems()（支持多选，从后往前删避免行号上移）。
; v6.8.23 1. 页眉“下划线”改为**段落的下边框**（v6.8.22 的 Font.Underline 只给文字
;           画线，文字短就只画一小截，不符合“横跨整个段落、左右撑满”的要求）：
;           通过 ParagraphFormat.Borders.Item(-3)（wdBorderBottom）设 LineStyle=1
;           单实线、LineWidth=6（0.75 磅）、Color=0 黑色；关掉时设 LineStyle=0。
;           段落左右缩进在前面已归零，所以线会自然铺满整个页眉版心宽度。
;           “清空”模式也顺手把下边框去掉，避免上一版留下的线残留；
;           “不改变”一个节都不碰。摘要行与日志文案不变；
;        2. 页眉“下划线”复选框由 x550 右移到 x580，与“加粗”复选框拉开距离
;           （两框之间约 38 像素），右端 656 仍在“距上边”标签（674）左侧；
;        3. 首页任务列表的“结果”列不再给成功项追加“（有提示）”后缀，
;           只保留“✔ 成功”和 PDF 导出失败时的“（PDF 失败）”；提示明细仍
;           写进日志说明列，状态栏与右下角提示也不变。
; =====================================================================

; v6.8.24 1. 页眉“下划线”选项改名为**“页眉横线”**：方案字段仍叫 HeaderUnderline
;           以兼容已有方案，但界面复选框、摘要行、日志与文档统一用 WPS 的叫法；
;           实现不变：横线仍由 ApplyHeaderBottomBorder() 画在段落下边框上，
;           横跨整个页眉版心、左右撑满，不给文字加字符下划线；
;        2. 复选框宽度由 76 调整为 88（四个汉字），右端仍在“距上边”标签左侧。
; v6.8.25 1. 首页“排版任务列表”第三列由完整文件路径改为**所在文件夹**：
;           第三列显示 `SplitPath()` 解析出的父目录；完整路径移到第 4 列隐藏，
;           双击打开、右键“打开所在目录”、批量排版按行取路径、去重等逻辑全部读第 4 列；
;        2. 页眉设置里“加粗”复选框紧跟“加粗：”三个字（x502），“页眉横线”复选框
;           回移到 x534，两个复选框仍保持 12 像素间距，不再在两个复选框之间留大块空白。
; v6.8.26 1. 页眉文字改为**默认不记住，可手动勾选“记住文字”**：方案字段 HeaderKeep
;           新增（0/1，随方案保存）；界面在“页眉文字”右侧加“记住文字”复选框
;           （x600，默认不勾选）。未勾选时：页眉文字不会写进顶层或方案注册表键，
;           不会在下次打开或右键/命令行批量里复用；
;        2. **启动时一次性清理旧版残留**：对所有方案逐一删除未启用记住文字时的
;           HeaderText（仅删除这个字段本身，不碰字体、横线等其它参数）；
;        3. 页眉文字输入框从 w350 缩到 w280，新增的复选框补回空间，行布局仍
;           以 x674 的“对齐方式:”为右边界，不压住右侧控件。
; v6.8.27 1. “记住文字”复选框的文字改为**红色**（`cFF0000`，与“偏离国标”提示同色），
;           提醒用户它会让页眉文字常驻方案；只改颜色，不改字号/字体/位置。
;        2. 保留用户手工微调的页眉文字行布局：输入框 `w250`，“记住文字”放在
;           `x570 y605 w80 h18`，右端 650，仍不压住 x674 的“对齐方式:”。
; v6.8.28 1. **正文/标题段落样式化应用**：排版时给段落应用 Word/WPS 内置段落样式，
;           在保持现有字体/段距等直接格式不变的前提下建立文档结构，便于大纲导航、
;           样式窗格识别，也为后续目录/提纲功能打基础：
;           · 全文基线先设为“正文”(wdStyleNormal, -1)，再覆盖本方案直接格式；
;           · 主标题/附件标题应用“标题”(wdStyleTitle, -63)；
;           · 标题区副标题尝试应用“副标题”(wdStyleSubtitle, -75)，不支持时静默降级；
;           · 独立一级~四级标题分别应用“标题 1~4”(wdStyleHeading1..4, -2~-5)；
;           · 标题与正文混排在同一段（超长标题按标点拆分）时，仅对混排段保持正文样式，
;             不把整段提升为标题样式，避免大纲误判；正文段/主送机关/附件标记等应用正文样式。
;        2. 新增 `TrySetParaStyle()`，所有样式赋值均 try 包裹，兼容 WPS/受保护文档/缺失样式。
; v6.8.29 1. **先调整内置样式格式，再应用到段落**：排版打开文档后，先通过
;           `AdjustBuiltinStyleSpacing(doc)` 把文档内置的“正文”“标题”“副标题”、
;           “标题 1~4”样式的段前、段后、行前、行后间距和行距调整为国标要求
;           （无段前段后、方案指定的固定行距），避免段落继承样式自带的 8 pt
;           段后距、单倍行距等非国标间距；随后给相应段落应用已调整好的样式，
;           正文/标题/标题1~4 段落的段间距与行距与方案保持一致。
; v6.8.32 1. **修复“标题”字体被样式覆盖**：`AdjustBuiltinStyleSpacing(doc)` 不再依赖全局 `ptMTitle/ptH1/...`
;           等可能未赋值的变量，而是在函数内部根据 `ActiveStyle` 重新用 `ParseFontSize()` 计算各级字号，
;           先把文档内置样式（正文/标题/副标题/标题1~4）的字体、字号、加粗、段距、行距调整为方案参数，
;           再把段落应用已调整好的样式，使标题字体等不会被样式自带的错误格式覆盖，也不会因全局变量
;           未赋值而抛出“This global variable has not been assigned a value”。
; v6.8.31 1. **先调整好样式的格式，再将段落应用样式**：`AdjustBuiltinStyleSpacing(doc)` 不再只调整段距/行距，
;           还会把文档内置的“正文”“标题”“副标题”“标题 1~4”样式本身的字体、字号、加粗
;           调整为方案参数（与方案里的 H1~H4、标题、副标题、正文字段一致）。这样后续给
;           段落应用样式时，段落会继承已经调整好的样式字体格式，不会带上 Word/WPS 样式
;           自带的错误字体、8 pt 段后距、单倍行距等。
; v6.8.34 1. **副标题与主送机关不缩进**：`AdjustBuiltinStyleSpacing(doc)` 中副标题（-75）的
;           `CharacterUnitFirstLineIndent` 设为 0，避免副标题被误加首行缩进；`GetOrCreateMainAddresseeStyle()`
;           在获取已存在的“主送机关”样式时也把首行缩进重置为 0，避免旧文档残留缩进。正文（-1）
;           与标题1~4（-2~-5）仍保持首行缩进 2 字符。
;        2. **主窗口页眉信息加粗**：`uiInfoHeader`（页眉处理方式摘要）设为粗体，更显著。
; v6.8.33 1. **标题与正文首行缩进统一**：`AdjustBuiltinStyleSpacing(doc)` 在调整内置样式时，
;           同时写入样式本身的首行缩进：正文（-1）及标题1~4（-2~-5）的 `CharacterUnitFirstLineIndent`
;           设为 2（2 字符），`FirstLineIndent` 设为 0；正文与标题套用样式后，首行缩进即由样式
;           继承，用户在 Word/WPS 样式框里点“正文”“标题 1”等时，首行缩进不会丢失。
; v6.8.30 1. **主送机关行单独建顶格段落样式**：新增 `GetOrCreateMainAddresseeStyle(doc)`，
;           在文档中查找或创建名为“主送机关”的自定义段落样式（WdStyleType=1 段落样式），
;           该样式本身即为顶格（Alignment:=0、首行缩进:=0、段前段后:=0）。排版时主送机关
;           行（文档正文第一段且以冒号结尾）直接套用该样式，再重新施加方案的正文字体/
;           字号/加粗等直接格式，避免每次都要在“正文”样式基础上反复覆盖间距，也便于
;           用户通过 Word/WPS 样式窗格统一管理此类段落。样式创建/应用全部 try 包裹，兼容
;           WPS/旧模板/受保护文档。
; =====================================================================
; --- 全局版本号定义 ---
global APP_VERSION   := "6.8.34"

; --- 默认国标字体名称预设 ---
global FONT_TITLE    := "方正小标宋简体" ; 国标字体名称预设（参考值，运行时以当前字体方案为准）
global FONT_SUBTITLE := "楷体"          ; 国标字体名称预设（参考值，运行时以当前字体方案为准）
global FONT_AUTHOR   := "楷体"          ; 国标字体名称预设（参考值，运行时以当前字体方案为准）
global FONT_H1       := "黑体"          ; 国标字体名称预设（参考值，运行时以当前字体方案为准）
global FONT_H2       := "楷体"          ; 国标字体名称预设（参考值，运行时以当前字体方案为准）
global FONT_BODY     := "仿宋"          ; 国标字体名称预设（参考值，运行时以当前字体方案为准）

; --- 运行时生效字号与行距变量 ---
global SIZE_TITLE        := 22     ; 字号参考值（运行时以当前字体方案经 ParseFontSize 转换后的磅值为准，二号=22磅）
global SIZE_BODY         := 16     ; 字号参考值（运行时以当前字体方案经 ParseFontSize 转换后的磅值为准，三号=16磅）
global SIZE_TABLE        := 10.5   ; 字号参考值（运行时以当前字体方案经 ParseFontSize 转换后的磅值为准，五号=10.5磅）
global LINE_SPACING_RULE := 4      ; Word行距规则：0=单倍, 4=固定值(磅), 5=多倍行距
global LINE_SPACING_BODY := 28     ; 正文默认行距：固定值 28 磅（单位：磅）
global LINE_SPACING_TITLE:= 35     ; 主标题默认行距：固定值 35 磅（单位：磅，运行时会按主标题字号×1.55 重算）

; --- 注册表存储路径定义（设置不再使用 config.ini） ---
global REG_PATH          := "HKEY_CURRENT_USER\Software\DocFormatTool"
global REG_SCHEMES_PATH  := "HKEY_CURRENT_USER\Software\DocFormatTool\FontSchemes"

; --- 页面版心边距参数（单位：毫米 mm） ---
global MarginType        := 1      ; 边距预设模式：1=党政机关公文格式, 2=WORD默认边距, 3=用户自定义
global MarginTop         := 37     ; 国标天头（上边距）37mm
global MarginBottom      := 35     ; 国标地脚（下边距）35mm
global MarginLeft        := 28     ; 国标订口（左边距）28mm
global MarginRight       := 26     ; 国标翻口（右边距）26mm
global CustomTop         := 0.0    ; 用户自定义上边距缓存
global CustomBottom      := 0.0    ; 用户自定义下边距缓存
global CustomLeft        := 0.0    ; 用户自定义左边距缓存
global CustomRight       := 0.0    ; 用户自定义右边距缓存

; --- 文档网格参数（v6.8.17：随方案保存，可在编辑方案窗口中修改） ---
; 对应 Word「页面设置 → 文档网格」里的三个选项，排版时由 ApplyDocumentGrid() 写入 COM。
global GridMode          := 1      ; 网格类型：0=不指定网格, 1=指定行和字符网格, 2=只指定行网格
global GridLines         := 22     ; 每页行数（GridMode 为 1 或 2 时生效）
global GridChars         := 28     ; 每行字符数（仅 GridMode=1 时生效；0=不指定）

; --- 页码格式参数 ---
global PageNumFormat     := "格式数字 — 1 — 居中"
global SkipSinglePageNum := 1      ; 单页文档是否不加页码

; --- 页眉参数（v6.8.20：随方案保存；本版只做界面与保存，排版时暂不写 Word 页眉）---
; 处理方式三态：0=不改变（原样保留文档原有页眉，默认）、1=添加·替换、2=清空。
; 之所以用 0/1/2 而不是下拉框下标：注册表和方案里存的是语义值，改下拉文案/顺序不影响旧数据。
global HeaderMode        := 0      ; 页眉处理方式：0=不改变, 1=添加·替换, 2=清空
global HeaderText        := ""     ; 页眉文字；空串=用文档文件名（不含扩展名），是合法值
; v6.8.26：页眉文字是否写入注册表（随方案保存）。0=不记住（默认），1=记住。
; 默认不勾选“记住文字”，保证下次打开/右键批量不会再显示上一次的单位名与主题。
global HeaderKeep        := 0
global HeaderAlign       := "居中" ; 页眉对齐：居左 / 居中 / 居右
global HeaderMarginTop   := 15     ; 页眉基线距页面上边（毫米），与 PageSetup.HeaderDistance 对应
global HeaderFont        := "宋体" ; 页眉字体（不进 GetFontKeys()，国标没有页眉字体规定）
global HeaderSize        := "四号" ; 页眉字号
global HeaderBold        := 0      ; 页眉是否加粗：0=不加粗, 1=加粗
global HeaderUnderline   := 0      ; 页眉是否加页眉横线（v6.8.22 起；v6.8.24 改名）：0=无, 1=有

; --- 当前选中的字体方案名称与自定义方案字典 ---
global CurrentFontScheme := "党政机关公文格式"
global CustomFontSchemes := Map()

; --- 排版附加功能开关配置 ---
; v6.8.18 起不再有“规范发文字号六角括号”开关（OptFixDocNum 已删除）
global ExportPDF        := 0       ; 是否同时导出 PDF 副本
global CreateBAK        := 1       ; 排版前是否自动生成原始备份
global OptPreventOrphan := 0       ; 是否启用标题防孤行
global OptAttachmentNewPage := 1   ; 是否让“附件1”标记另起一页（GB/T 9704-2012：附件另面编排）
global gAttachSectionIndex := 0    ; v6.8.15：附件另起一页时，附件所在节的节号，用于在 SetPageNumbers 强制页码从 1 开始
global wdSectionBreakNextPage := 2 ; Word WdSectionBreak 枚举：wdSectionBreakNextPage = 2（在当前位置下一页开始新节）
global MaxHeadingLen    := 35      ; 标题最大字符数拦截阈值
global ContextActive    := 0       ; 用户是否主动添加过右键菜单标记

; --- 备份文件存储路径 ---
global DefaultBakDir := A_MyDocuments "\公文自动排版文件备份"
global BakDir        := DefaultBakDir

; --- 右下角无焦点悬浮通知框（Toast）全局控件句柄 ---
global toastGui       := ""
global toastTitleCtrl := ""
global toastMsgCtrl   := ""
global toastIconCtrl  := ""

; --- “打开备份目录”方形按钮所用系统图标句柄（需常驻，不能提前释放） ---
global BakOpenHIcon := 0

; --- 资源管理器右键多选并发队列与进程锁文件路径 ---
global QueueFilePath := A_Temp "\DocFormatTool_TaskQueue.txt"
global QueueLockPath := A_Temp "\DocFormatTool_Queue.lock"

; --- 本程序创建的后台 Word/WPS 进程 PID 追踪表 ---
global ManagedWordPIDs := Map()

; --- 注册表配置方案存储路径 ---
global REG_PROFILES_PATH := REG_PATH "\Profiles"
global BUILTIN_PROFILE_NAME := "党政机关公文格式"

; --- 主界面 GUI 与各控件全局引用 ---
global isProgramUpdating := false
; v6.8.22：GUI 批量排版是否正在运行。列表右键菜单不归 SetGUIControlsLock() 管，
; 靠这个标志拦住“排版途中移除/清空/恢复行”——那会把 Loop count 的行号搞错。
global GuiBatchRunning := false
global mainGui, uiProfile, uiBtnEditProfile, uiBtnSaveAsProfile
global uiInfoMarginType, uiInfoMargins, uiInfoPage, uiInfoSkip, uiInfoFont, uiInfoLen, uiInfoOptions, uiInfoHeader
global uiPDF, uiBAK, uiBakPath, uiBakBtn, uiBakOpen, uiLV, uiBtnRun, uiBtnClear, uiBtnKill, uiBtnExport, uiBtnLog, uiSB
global uiSettingsMenu

; --- 当前配置方案名称 ---
global ActiveProfile := BUILTIN_PROFILE_NAME

; --- 方案参数编辑窗口及控件引用（窗口关闭后置空） ---
global editGui := "", dlgName, dlgMarginType, dlgTop, dlgBottom, dlgLeft, dlgRight
global dlgTopUD, dlgBottomUD, dlgLeftUD, dlgRightUD
global dlgFontScheme, dlgBtnManageScheme, dlgPageFormat, dlgSkipSingle
global dlgProtectTable, dlgPreventOrphan, dlgAttachNewPage, dlgMaxHeadingLen
global dlgCustomTop, dlgCustomBottom, dlgCustomLeft, dlgCustomRight
; 页眉设置组控件（v6.8.20）：处理方式、文字、对齐、字体、字号、加粗、距上边 + 微调按钮
global dlgHeaderMode, dlgHeaderText, dlgHeaderAlign
global dlgHeaderFont, dlgHeaderSize, dlgHeaderBold, dlgHeaderUnderline
global dlgHeaderMargin, dlgHeaderMarginUD, dlgHeaderKeep
; 编辑窗口内的字体字号行控件、行距控件及红色星号
global dlgFontRows := [], dlgSpacingCtrl := "", dlgSpacingStar := "", dlgSpacingLabel := ""
; 当前正在显示悬停说明的选项编号（0 表示未显示；1/2/3 对应三个排版选项）
global optTipId := 0
; 打开编辑窗口时的原方案名，用于保存时判断是否需要重命名/删除旧键
global editorOriginalName := ""

; 字体方案管理窗口的行名称与控件对照表（用于国标偏差星号）
global fontMgrRows := [], fontMgrSpacingCtrl := "", fontMgrSpacingLabel := "", fontMgrSpacingStar := ""

; =====================================================================
; [系统字体枚举核心函数：从注册表安全读取本机已安装的中文字体]
; =====================================================================
GetSystemChineseFonts() {
    ; 定义默认的常用中文字体列表
    fontList := ["方正小标宋简体", "仿宋", "黑体", "楷体", "宋体"]
    fontMap := Map()
    ; 遍历默认字体列表并写入 Map
    for f in fontList
        fontMap[f] := true

    ; 定义系统注册表中的字体存放路径数组
    regKeys := [
        "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
        "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    ]

    ; 循环读取注册表项中的所有字体键值
    for rootKey in regKeys {
        Loop Reg, rootKey, "V" {
            ; 清理并去除字体名称后面的括号修饰
            cleanName := Trim(RegExReplace(A_LoopRegName, "\s*\([^)]*\)$", ""))
            ; 排除空名称和竖排字体，并且在 Map 中不存在时处理
            if (cleanName != "" && SubStr(cleanName, 1, 1) != "@" && !fontMap.Has(cleanName)) {
                ; 仅匹配包含中文字符的字体名称
                if RegExMatch(cleanName, "[\x{4e00}-\x{9fa5}]") {
                    fontMap[cleanName] := true
                    fontList.Push(cleanName)
                }
            }
        }
    }
    ; 返回最终收集到的字体数组
    return fontList
}

; 初始化全局字体列表
global WpsFontNameList := GetSystemChineseFonts()

; 初始化常用字号列表
global WpsFontSizeList := [
    "初号", "小初", "一号", "小一", "二号", "小二",
    "三号", "小三", "四号", "小四", "五号", "小五",
    "六号", "小六", "七号", "八号",
    "42", "36", "26", "24", "22", "18", "16", "15", "14", "12", "10.5", "9", "7.5"
]

; 初始化行距预设列表
global LineSpacingPresets := [
    "固定值 28 磅", "固定值 30 磅", "单倍行距", "1.5 倍行距", "2 倍行距"
]

; 初始化页码格式预设列表
global PageFormatList := [
    "无",
    "普通数字 1     居左", "普通数字 1     居中", "普通数字 1     居右",
    "格式数字 — 1 — 居左", "格式数字 — 1 — 居中", "格式数字 — 1 — 居右",
    "加总页码 1 / n 居左", "加总页码 1 / n 居中", "加总页码 1 / n 居右"
]

; =====================================================================
; [默认国家标准公文方案参数模型 (GB/T 9704-2012)]
; =====================================================================
GetDefaultSchemeMap() {
    ; 返回包含公文各部件默认字体、字号、加粗属性的对象
    return {
        MTitleFont:   "方正小标宋简体", MTitleSize:   "二号", MTitleBold:   0,
        SubTitleFont: "楷体",           SubTitleSize: "三号", SubTitleBold: 0,
        AuthorFont:   "楷体",           AuthorSize:   "三号", AuthorBold:   0,
        BodyFont:     "仿宋",           BodySize:     "三号", BodyBold:     0,
        H1Font:       "黑体",           H1Size:       "三号", H1Bold:       0,
        H2Font:       "楷体",           H2Size:       "三号", H2Bold:       0,
        H3Font:       "仿宋",           H3Size:       "三号", H3Bold:       1,
        H4Font:       "仿宋",           H4Size:       "三号", H4Bold:       0,
        TableFont:    "仿宋",           TableSize:    "五号", TableBold:    0,
        ; 附件标记（“附件1”“附件 1”这一行）：GB/T 9704-2012 7.3.8 规定 3 号黑体、顶格
        AttachMarkFont: "黑体",         AttachMarkSize: "三号", AttachMarkBold: 0,
        ; 页码：GB/T 9704-2012 7.3.7 规定用 4 号半角宋体阿拉伯数字
        PageNumFont:  "宋体",           PageNumSize:  "四号", PageNumBold:  0,
        Spacing:      "固定值 28 磅"
    }
}

; 初始化当前激活的样式对象
global ActiveStyle := GetDefaultSchemeMap()

; 方案中随整体一起保存的字体字号字段
GetFontKeys() {
    return ["MTitleFont", "MTitleSize", "MTitleBold",
            "SubTitleFont", "SubTitleSize", "SubTitleBold",
            "AuthorFont", "AuthorSize", "AuthorBold",
            "BodyFont", "BodySize", "BodyBold",
            "H1Font", "H1Size", "H1Bold",
            "H2Font", "H2Size", "H2Bold",
            "H3Font", "H3Size", "H3Bold",
            "H4Font", "H4Size", "H4Bold",
            "TableFont", "TableSize", "TableBold",
            "AttachMarkFont", "AttachMarkSize", "AttachMarkBold",
            "PageNumFont", "PageNumSize", "PageNumBold",
            "Spacing"]
}

; 编辑窗口中字体字号各行的定义（名称、字段前缀、国标字体/字号/加粗、国标说明）
GetFontItemsDef() {
    return [
        ["主标题:",   "MTitle",   "方正小标宋简体", "二号", 0, "小标宋 二号 不加粗"],
        ["副标题:",   "SubTitle", "楷体",           "三号", 0, "楷体 三号 不加粗"],
        ["作者/发文:", "Author",   "楷体",           "三号", 0, "楷体 三号 不加粗"],
        ["正文内容:", "Body",     "仿宋",           "三号", 0, "仿宋 三号 不加粗"],
        ["一级标题:", "H1",       "黑体",           "三号", 0, "黑体 三号 不加粗"],
        ["二级标题:", "H2",       "楷体",           "三号", 0, "楷体 三号 不加粗"],
        ["三级标题:", "H3",       "仿宋",           "三号", 1, "仿宋 三号 (常加粗)"],
        ["四级标题:", "H4",       "仿宋",           "三号", 0, "仿宋 三号 不加粗"],
        ["表格内容:", "Table",    "仿宋",           "五号", 0, "仿宋 五号 不加粗"],
        ; v6.8.14：附件标记移到表格之后（窗口最下面），其下新增“页码”行
        ["附件标记:", "AttachMark", "黑体",         "三号", 0, "黑体 三号 不加粗"],
        ["页码:",     "PageNum",  "宋体",           "四号", 0, "宋体 四号 不加粗"]
    ]
}

; 把字体字号对象装载为运行时样式（方案内的字体参数优先）
ApplyStyleObject(src) {
    global ActiveStyle, LINE_SPACING_RULE, LINE_SPACING_BODY, LINE_SPACING_TITLE
    def := GetDefaultSchemeMap()
    st := {}
    for k, defaultV in def.OwnProps() {
        v := defaultV
        if (IsObject(src) && src.HasOwnProp(k) && String(src.%k%) != "")
            v := src.%k%
        if (InStr(k, "Bold"))
            v := Integer(v)
        st.%k% := v
    }
    ActiveStyle := st

    ParseLineSpacing(st.Spacing, &rule, &val)
    LINE_SPACING_RULE := rule
    LINE_SPACING_BODY := val
    titlePt := ParseFontSize(st.MTitleSize, 22.0)
    LINE_SPACING_TITLE := Round(titlePt * 1.55)
}

; =====================================================================
; [配置持久化读写抽象层]
; =====================================================================
GetConfigVal(section, key, defaultVal) {
    global REG_PATH
    ; 6.8.0 起设置只存放在注册表；section 参数仅为兼容既有调用而保留
    try return RegRead(REG_PATH, key)
    catch
        return defaultVal
}

SetConfigVal(section, key, val) {
    global REG_PATH
    ; 只写入注册表，不再创建或同步 config.ini
    try RegWrite(val, "REG_SZ", REG_PATH, key)
}

; =====================================================================
; [自定义字体方案加载、保存与删除函数]
; =====================================================================
LoadFontSchemes() {
    global CustomFontSchemes, REG_SCHEMES_PATH
    CustomFontSchemes := Map()
    def := GetDefaultSchemeMap()

    ; 自定义字体方案只从注册表读取
    Loop Reg, REG_SCHEMES_PATH, "K" {
        sName := A_LoopRegName
        try {
            item := {}
            for k, defaultV in def.OwnProps() {
                val := RegRead(REG_SCHEMES_PATH "\" sName, k)
                item.%k% := InStr(k, "Bold") ? (IsNumber(val) ? Integer(val) : 0) : val
            }
            CustomFontSchemes[sName] := item
        }
    }
}

SaveFontSchemeData(sName, schemeObj) {
    global REG_SCHEMES_PATH
    ; 自定义字体方案只写入注册表
    try {
        subKey := REG_SCHEMES_PATH "\" sName
        for k, v in schemeObj.OwnProps()
            RegWrite(v, "REG_SZ", subKey, k)
    }
}

DeleteFontSchemeData(sName) {
    global REG_SCHEMES_PATH
    try RegDeleteKey(REG_SCHEMES_PATH "\" sName)
}

ExportSettingsToReg() {
    global REG_PATH, uiSB

    ; 先把当前界面参数写入注册表，确保导出的内容与界面一致
    SaveAllCurrentSettings()

    ; 默认保存到程序所在目录：编译后为 EXE 所在目录，未编译时为脚本所在目录
    defaultFile := A_ScriptDir "\公文排版设置.reg"
    regFile := FileSelect("S16", defaultFile, "导出设置（注册表文件）", "注册表文件 (*.reg)")
    if (regFile == "")
        return
    if (SubStr(regFile, -4) != ".reg")
        regFile .= ".reg"

    ; 直接导出本程序的注册表根键，包含全部设置和自定义字体方案
    try
        exitCode := RunWait('reg.exe export "' REG_PATH '" "' regFile '" /y', , "Hide")
    catch {
        ShowFloatingToast("导出失败", "无法调用系统注册表导出工具 reg.exe。", 3000, false)
        return
    }

    if (exitCode != 0 || !FileExist(regFile)) {
        ShowFloatingToast("导出失败", "注册表导出失败，请检查目标路径的写入权限！", 3000, false)
        uiSB.SetText(" 导出失败：无法写入 " regFile)
        return
    }

    uiSB.SetText(" 设置已导出：" regFile)
    ; 用右下角浮动提示替代模态对话框，不打断当前操作；完整路径同时显示在状态栏
    ShowFloatingToast("导出成功", "注册表文件已保存：`n" regFile "`n换电脑双击该文件导入，重新启动即可恢复设置。", 5000, true)
}

ParseFontSize(inputStr, defaultVal := 16.0) {
    str := Trim(inputStr)
    nameMap := Map(
        "初号", 42.0, "小初", 36.0, "一号", 26.0, "小一", 24.0,
        "二号", 22.0, "小二", 18.0, "三号", 16.0, "小三", 15.0,
        "四号", 14.0, "小四", 12.0, "五号", 10.5, "小五", 9.0,
        "六号", 7.5,  "小六", 6.5,  "七号", 5.5,  "八号", 5.0
    )
    if nameMap.Has(str)
        return nameMap[str]
    numOnly := RegExReplace(str, "[^\d\.]", "")
    return (IsNumber(numOnly) && Float(numOnly) > 0) ? Float(numOnly) : defaultVal
}

ParseLineSpacing(inputStr, &ruleOut, &valOut) {
    str := Trim(inputStr)
    numPart := RegExReplace(str, "[^\d\.]", "")
    valNum := (IsNumber(numPart) && Float(numPart) > 0) ? Float(numPart) : 28.0

    if InStr(str, "倍") {
        if (valNum == 1.0 || valNum == 28.0) {
            ; 兼容 "单倍行距"（无数字）与旧版 "1 倍行距"
            ruleOut := 0, valOut := 12
        } else {
            ruleOut := 5, valOut := valNum * 12
        }
    } else {
        ruleOut := 4, valOut := valNum
    }
}

ApplyFontScheme(schemeName) {
    global CustomFontSchemes
    ; 旧接口：仅用于启动时兼容历史字体方案，新方案不再使用独立字体方案
    if (schemeName == GetBuiltinProfileName() || !CustomFontSchemes.Has(schemeName))
        ApplyStyleObject(GetDefaultSchemeMap())
    else
        ApplyStyleObject(CustomFontSchemes[schemeName])
}

; =====================================================================
; [程序主入口]
; =====================================================================
LoadSettings()
; 读取/迁移整体配置方案，并把当前方案装载为生效参数
EnsureProfiles()

; 判断是否为编译后运行且用户主动添加过右键菜单，若路径失效则自愈
if (A_IsCompiled && ContextActive == 1) {
    CheckAndRepairContextMenuPath()
}

; 检查是否存在命令行参数
if (A_Args.Length > 0) {
    AppendTasksToQueue(A_Args)
    if (!TryAcquireMasterLock())
        ExitApp()

    Sleep(350)
    task_files := ReadAndClearQueue()
    ReleaseMasterLock()

    if (task_files.Length > 0)
        RunSilentProcessing(task_files)
    else
        ExitApp()
} else {
    CreateTraySettingsMenu()
    ShowGUI()
}
return

; =====================================================================
; [多文件并发排队队列管理函数]
; =====================================================================
AppendTasksToQueue(args) {
    Loop 20 {
        try {
            f := FileOpen(QueueFilePath, "a", "UTF-8")
            if (f) {
                for arg in args {
                    if (arg != "")
                        f.WriteLine(arg)
                }
                f.Close()
                return true
            }
        } catch {
            Sleep(25)
        }
    }
    return false
}

TryAcquireMasterLock() {
    try {
        if (FileExist(QueueLockPath)) {
            fileTime := FileGetTime(QueueLockPath, "M")
            if (DateDiff(A_Now, fileTime, "Seconds") > 30)
                FileDelete(QueueLockPath)
            else
                return false
        }
        FileAppend(A_Now "`n", QueueLockPath, "UTF-8")
        return true
    } catch {
        return false
    }
}

ReleaseMasterLock() {
    try {
        if (FileExist(QueueLockPath))
            FileDelete(QueueLockPath)
    }
}

ReadAndClearQueue() {
    taskList := []
    uniqueMap := Map()
    if (!FileExist(QueueFilePath))
        return taskList

    Loop 20 {
        try {
            content := FileRead(QueueFilePath, "UTF-8")
            FileDelete(QueueFilePath)
            Loop Parse, content, "`n", "`r" {
                itemPath := Trim(A_LoopField)
                if (itemPath != "" && !uniqueMap.Has(itemPath)) {
                    uniqueMap[itemPath] := true
                    taskList.Push(itemPath)
                }
            }
            break
        } catch {
            Sleep(25)
        }
    }
    return taskList
}

; =====================================================================
; [备份回滚与恢复机制：查找指定文件最近的备份并恢复]
; =====================================================================
RestoreLatestBackup(filePath) {
    global BakDir, DefaultBakDir
    targetDir := (BakDir != "") ? BakDir : DefaultBakDir
    if (!DirExist(targetDir) || !FileExist(filePath)) {
        ShowFloatingToast("恢复失败", "未找到文件或备份目录不存在！", 3000, false)
        return false
    }

    SplitPath(filePath, , , &fileExt, &nameNoExt)
    latestBak := "", latestTime := ""

    Loop Files, targetDir "\" nameNoExt "_*." fileExt {
        if (latestTime == "" || A_LoopFileTimeModified > latestTime) {
            latestTime := A_LoopFileTimeModified
            latestBak := A_LoopFilePath
        }
    }

    if (latestBak != "") {
        try {
            FileCopy(latestBak, filePath, true)
            ShowFloatingToast("恢复成功", "已将文档恢复至最近一次排版前的备份！", 3000, true)
            return true
        } catch {
            ShowFloatingToast("恢复失败", "文件被占用，请先关闭相关 Word 窗口！", 3500, false)
            return false
        }
    } else {
        ShowFloatingToast("恢复失败", "在备份目录中未找到此文件的历史备份！", 3000, false)
        return false
    }
}

; =====================================================================
; [核心界面 (GUI) 构建与事件绑定]
; =====================================================================
ShowGUI() {
    global APP_VERSION
    global mainGui, uiProfile, uiBtnEditProfile, uiBtnSaveAsProfile
    global uiInfoMarginType, uiInfoMargins, uiInfoPage, uiInfoSkip, uiInfoFont, uiInfoLen, uiInfoOptions, uiInfoHeader
    global uiPDF, uiBAK, uiBakPath, uiBakBtn, uiBakOpen, uiLV, uiBtnRun, uiBtnClear, uiBtnKill, uiBtnExport, uiBtnLog, uiSB
    global uiSettingsMenu
    global ExportPDF, CreateBAK, BakDir, ActiveProfile

    if (IsSet(mainGui) && IsObject(mainGui)) {
        try {
            mainGui.Show()
            return
        }
    }

    mainGui := Gui("+Theme -MaximizeBox", "党政机关公文自动排版引擎 V" APP_VERSION)
    mainGui.OnEvent("Close", GuiClose)
    mainGui.OnEvent("DropFiles", GuiDropFiles)
    mainGui.SetFont("s10", "Microsoft YaHei")
    mainGui.BackColor := "FFFFFF"

    ; --- 顶部菜单栏（备份开关与标题阈值已移出菜单）---
    mainMenu := MenuBar()
    settingsMenu := Menu()
    settingsMenu.Add("添加 Word 文件右键运行命令", (*) => RegisterDocumentContextMenu(true))
    settingsMenu.Add("删除 Word 文件右键运行命令", (*) => RemoveDocumentContextMenu())
    settingsMenu.Add()
    settingsMenu.Add("导出当前所有设置 (注册表 .reg)", (*) => ExportSettingsToReg())
    mainMenu.Add("设置(&S)", settingsMenu)

    helpMenu := Menu()
    helpMenu.Add("请作者喝一杯咖啡(&A)", (*) => ShowAboutDialog())
    mainMenu.Add("帮助(&H)", helpMenu)
    mainGui.MenuBar := mainMenu

    ; 记录菜单对象，供批量排版期间锁定使用
    uiSettingsMenu := settingsMenu

    ; --- 顶部主横幅 ---
    mainGui.SetFont("s18 bold c0055AA")
    mainGui.Add("Text", "x20 y15 w640 Center BackgroundTrans", "党政机关公文自动排版引擎")
    mainGui.SetFont("s9 norm c666666")
    mainGui.Add("Text", "x20 y45 w640 Center BackgroundTrans", " (符合 GB/T 9704-2012 国标规范)  |  请将 Word 文件拖入下方列表")
    mainGui.SetFont("s10 c000000")

    ; --- 配置方案选择区 ---
    ; v6.8.20：组框由 h165 加高到 h190，在第四行“排版选项”下面追加第五行“页眉处理”。
    ; 追加在末尾而不是插在中间，已有四行的 y 坐标一个都不用动。
    mainGui.Add("GroupBox", "x20 y75 w640 h190", "配置方案（单位/文种）")
    mainGui.Add("Text", "x35 y98 w60", "当前方案:")
    uiProfile := mainGui.Add("DropDownList", "x100 y94 w300")
    uiProfile.OnEvent("Change", OnProfileChange)

    uiBtnEditProfile := mainGui.Add("Button", "x410 y93 w90 h28", "编辑参数…")
    uiBtnEditProfile.OnEvent("Click", OpenProfileEditor)

    uiBtnSaveAsProfile := mainGui.Add("Button", "x508 y93 w80 h28", "另存为…")
    uiBtnSaveAsProfile.OnEvent("Click", SaveProfileAs)

    ; 当前方案信息（只读摘要）
    mainGui.SetFont("s9 c808080")
    mainGui.Add("Text", "x40 y137 w55", "边距预设")
    mainGui.Add("Text", "x315 y137 w70", "上/下/左/右")
    mainGui.Add("Text", "x40 y162 w55", "页码格式")
    mainGui.Add("Text", "x330 y162 w55", "单页不加")
    mainGui.Add("Text", "x40 y187 w55", "字体字号")
    mainGui.Add("Text", "x330 y187 w55", "标题阈值")
    mainGui.Add("Text", "x40 y212 w55", "排版选项")
    mainGui.Add("Text", "x40 y237 w55", "页眉处理")
    mainGui.SetFont("s9 c000000")
    uiInfoMarginType := mainGui.Add("Text", "x100 y137 w205")
    uiInfoMargins    := mainGui.Add("Text", "x390 y137 w250")
    uiInfoPage       := mainGui.Add("Text", "x100 y162 w215")
    uiInfoSkip       := mainGui.Add("Text", "x390 y162 w150")
    uiInfoFont       := mainGui.Add("Text", "x100 y187 w215")
    uiInfoLen        := mainGui.Add("Text", "x390 y187 w150")
    uiInfoOptions    := mainGui.Add("Text", "x100 y212 w540")
    uiInfoHeader     := mainGui.Add("Text", "x100 y237 w540")

    ; --- 输出与备份（全局，不随方案切换）---
    ; v6.8.20：以下四块整体下移 25 像素，给“配置方案”组新增的第五行摘要让位。
    mainGui.SetFont("s10 c000000")
    mainGui.Add("GroupBox", "x20 y275 w640 h90", "输出与备份（不随方案切换）")
    uiPDF := mainGui.Add("Checkbox", "x35 y300 w200 Checked" ExportPDF, "同时生成 PDF 副本")
    uiBAK := mainGui.Add("Checkbox", "x255 y300 w220 Checked" CreateBAK, "排版前生成原始备份")
    mainGui.SetFont("s8 c909090")
    mainGui.Add("Text", "x485 y304 w160", "切换方案不会改变这三项")
    mainGui.SetFont("s10 c000000")

    uiPDF.OnEvent("Click", (ctrl, *) => SaveSingleSetting("Options", "ExportPDF", ctrl.Value))
    uiBAK.OnEvent("Click", (ctrl, *) => OnBackupCheckboxChange(ctrl))

    mainGui.Add("Text", "x35 y335 w65", "备份目录:")
    uiBakPath := mainGui.Add("Edit", "x105 y331 w400 h27", BakDir)
    uiBakBtn := mainGui.Add("Button", "x510 y331 w80 h27", "浏览...")
    ; 正方形图标按钮：BS_ICON(0x40) + BM_SETIMAGE 只显示图标，不绘制文字
    uiBakOpen := mainGui.Add("Button", "x595 y331 w27 h27 +0x40", "打开")
    SetButtonStockIcon(uiBakOpen, 4) ; 4 = SIID_FOLDEROPEN（系统标准“打开文件夹”图标）

    uiBakPath.OnEvent("Change", (ctrl, *) => SaveSingleSetting("Options", "BakDir", ctrl.Value))
    uiBakBtn.OnEvent("Click", SelectBakPath)
    uiBakOpen.OnEvent("Click", OpenBakFolder)

    ; --- 下部：拖拽任务列表区 ---
    mainGui.Add("GroupBox", "x20 y375 w640 h165", "排版任务列表 (拖拽文件、文件夹至此)")
    ; v6.8.25：第三列显示“所在文件夹”（父目录），更短更容易一眼看全；
    ; 完整路径移到第 4 列并宽度设为 0 隐藏，所有读取路径的地方都改用第 4 列。
    uiLV := mainGui.Add("ListView", "x30 y400 w620 h130 Grid", ["文件名", "状态", "所在文件夹", "文件路径"])
    uiLV.ModifyCol(1, 180)
    uiLV.ModifyCol(2, 130)
    uiLV.ModifyCol(3, 280)
    uiLV.ModifyCol(4, 0)

    ; 绑定任务列表右键菜单（v6.8.22 参照 PDF 加页码工具补齐）：
    ; 打开文件 / 打开所在目录 / 分隔线 / 从备份恢复此文件 / 分隔线 /
    ; 移除所选文件 / 清空列表。只在 item > 0（确实点在某一行上）时弹出，
    ; 空白处右键不弹，避免用户以为点了空白也能操作。
    lvMenu := Menu()
    lvMenu.Add("打开文件", (*) => OpenSelectedLVFile())
    lvMenu.Add("打开所在目录", (*) => OpenSelectedLVFolder())
    lvMenu.Add()                       ; 分隔线
    lvMenu.Add("从备份恢复此文件", (*) => RestoreSelectedLVItem())
    lvMenu.Add()
    lvMenu.Add("移除所选文件", (*) => RemoveSelectedLVItems())
    lvMenu.Add("清空列表", (*) => ClearList())
    uiLV.OnEvent("ContextMenu", (ctrl, item, isRightClick, x, y) => (item > 0 ? lvMenu.Show() : ""))
    ; 双击任务列表中的文件，用系统默认程序打开，便于排版前核对原文件
    uiLV.OnEvent("DoubleClick", (*) => OpenSelectedLVFile())

    ; --- 底部：功能操作按钮栏 ---
    ; 五个控件排成一行：清空列表、导出设置、日志、结束后台排版进程、
    ; 保存参数并运行。总宽度 640，间隔 8 像素。
    uiBtnClear := mainGui.Add("Button", "x20 y550 w102 h35", "清空列表")
    uiBtnClear.OnEvent("Click", ClearList)

    uiBtnExport := mainGui.Add("Button", "x130 y550 w102 h35", "导出设置")
    uiBtnExport.OnEvent("Click", (*) => ExportSettingsToReg())

    ; 日志按钮使用文字命名，避免纯图标看不出功能；点击后用默认程序打开执行日志
    uiBtnLog := mainGui.Add("Button", "x240 y550 w44 h35", "日志")
    uiBtnLog.OnEvent("Click", OpenLogFile)

    uiBtnKill := mainGui.Add("Text", "x292 y550 w150 h35 Center +0x200 BackgroundFFF9C4 Border c554400", "结束后台排版进程")
    uiBtnKill.SetFont("s10 bold")
    uiBtnKill.OnEvent("Click", KillAllWordProcesses)

    OnMessage(0x0200, CheckButtonHover)

    uiBtnRun := mainGui.Add("Text", "x450 y550 w210 h35 Center +0x200 BackgroundE8F5E9 Border c117711", "保存参数并运行  ▶")
    uiBtnRun.SetFont("s11 bold")
    uiBtnRun.OnEvent("Click", RunGUIProcessing)

    uiSB := mainGui.Add("StatusBar",, " 准备就绪。")

    mainGui.Show("w680 h625")

    ; 方案下拉与信息面板必须在控件创建完成后填充
    RefreshProfileDropdown()
    RefreshProfileInfo()
    if (IsSet(uiSB) && IsObject(uiSB))
        uiSB.SetText(" 当前方案：" ActiveProfile)
}

SaveSingleSetting(section, key, val) {
    SetConfigVal(section, key, val)
}

; =====================================================================
; [整体配置方案：注册表读写、切换与首页交互]
; =====================================================================
GetBuiltinProfileName() {
    global BUILTIN_PROFILE_NAME
    return BUILTIN_PROFILE_NAME
}

IsBuiltinProfile(name) {
    return (name == GetBuiltinProfileName())
}

; 内置只读方案“党政机关公文格式”的参数（GB/T 9704-2012）
GetBuiltinProfile() {
    p := {
        MarginType: 1,
        MarginTop: 37, MarginBottom: 35, MarginLeft: 28, MarginRight: 26,
        CustomTop: 0.0, CustomBottom: 0.0, CustomLeft: 0.0, CustomRight: 0.0,
        PageNumFormat: "格式数字 — 1 — 居中",
        SkipSinglePageNum: 1,
        ; 文档网格：国标公文版心为每页 22 行、每行 28 个汉字
        GridMode: 1, GridLines: 22, GridChars: 28,
        CurrentFontScheme: "党政机关公文格式",
        OptProtectTable: 0, OptPreventOrphan: 0,
        OptAttachmentNewPage: 1,
        MaxHeadingLen: 35,
        ; v6.8.20 页眉：默认“不改变”，即完全不碰文档原有页眉；页眉文字默认留空（= 用文件名）
        HeaderMode: 0, HeaderText: "", HeaderAlign: "居中",
        HeaderMarginTop: 15, HeaderFont: "宋体", HeaderSize: "四号", HeaderBold: 0,
        HeaderUnderline: 0, HeaderKeep: 0
    }
    ; 字体字号属于方案本身，与边距等参数一起保存
    for k, v in GetDefaultSchemeMap().OwnProps()
        p.%k% := v
    return p
}

; 方案包含的全部键（PDF、备份开关、备份目录不在此列）
GetProfileKeys() {
    keys := ["MarginType", "MarginTop", "MarginBottom", "MarginLeft", "MarginRight",
             "CustomTop", "CustomBottom", "CustomLeft", "CustomRight",
             "PageNumFormat", "SkipSinglePageNum", "CurrentFontScheme",
             "GridMode", "GridLines", "GridChars",
             "OptProtectTable", "OptPreventOrphan",
             "OptAttachmentNewPage", "MaxHeadingLen",
             "HeaderMode", "HeaderText", "HeaderAlign",
             "HeaderMarginTop", "HeaderFont", "HeaderSize", "HeaderBold",
             "HeaderUnderline", "HeaderKeep"]
    for k in GetFontKeys()
        keys.Push(k)
    return keys
}

GetUserProfileNames() {
    global REG_PROFILES_PATH
    names := []
    try {
        Loop Reg, REG_PROFILES_PATH, "K" {
            n := A_LoopRegName
            if (n != "" && !IsBuiltinProfile(n))
                names.Push(n)
        }
    }
    return names
}

ProfileExists(name) {
    global REG_PROFILES_PATH
    if (name == "")
        return false
    if (IsBuiltinProfile(name))
        return true
    try {
        Loop Reg, REG_PROFILES_PATH, "K" {
            if (A_LoopRegName == name)
                return true
        }
    }
    return false
}

; 读取方案参数；内置方案直接返回国标默认值，缺失的方案也回退到内置值
ProfileRead(name) {
    global REG_PROFILES_PATH, CustomFontSchemes
    def := GetBuiltinProfile()
    if (IsBuiltinProfile(name))
        return def
    if (!ProfileExists(name))
        return def

    ; 6.8.8 之前字体字号单独存为“字体方案”，这里按旧方案名称自动迁移为方案自己的字段
    legacyMap := ""
    try {
        legacyName := RegRead(REG_PROFILES_PATH "\" name, "CurrentFontScheme")
        if (legacyName != "" && !IsBuiltinProfile(legacyName) && CustomFontSchemes.Has(legacyName))
            legacyMap := CustomFontSchemes[legacyName]
    }

    obj := {}
    gridKeyMissing := false      ; v6.8.17：旧方案没有文档网格字段时，需要按边距预设推算
    for k in GetProfileKeys() {
        defaultV := def.%k%
        v := defaultV
        try {
            v := RegRead(REG_PROFILES_PATH "\" name, k)
        } catch {
            if (k == "GridMode")
                gridKeyMissing := true
            if (IsObject(legacyMap) && legacyMap.HasOwnProp(k))
                v := legacyMap.%k%
        }
        if (IsNumber(defaultV)) {
            if (InStr(String(v), "."))
                v := Float(v)
            else
                v := Integer(v)
        }
        obj.%k% := v
    }
    ; 6.8.16 及以前保存的方案里没有 GridMode/GridLines/GridChars 三个键。
    ; 这时按边距预设推算默认值，保证升级后行为与 6.8.16 一致：
    ;   边距预设为“党政机关公文格式”→ 指定行和字符网格、22 行、28 字（即内置默认值，无需改）
    ;   其他预设（WORD 默认边距 / 用户自定义）→ 只指定行网格、44 行、不指定每行字数
    if (gridKeyMissing && Integer(obj.MarginType) != 1) {
        obj.GridMode := 2
        obj.GridLines := 44
        obj.GridChars := 0
    }
    return obj
}

; 把方案对象写入注册表（方案名作为子键）
ProfileWriteObject(name, obj) {
    global REG_PROFILES_PATH
    try {
        subKey := REG_PROFILES_PATH "\" name
        for k in GetProfileKeys() {
            v := obj.%k%
            if (k == "MarginType" || k == "SkipSinglePageNum"
                || k == "OptProtectTable" || k == "OptPreventOrphan" || k == "OptAttachmentNewPage"
                || k == "MaxHeadingLen" || k == "GridMode" || k == "GridLines" || k == "GridChars"
                || k == "HeaderMode" || k == "HeaderUnderline" || k == "HeaderKeep" || InStr(k, "Bold"))
                v := Integer(v)
            else if (k == "MarginTop" || k == "MarginBottom" || k == "MarginLeft"
                || k == "MarginRight" || k == "HeaderMarginTop" || SubStr(k, 1, 6) == "Custom")
                v := FmtMargin(v)
            else
                v := String(v)
            RegWrite(v, "REG_SZ", subKey, k)
        }
    }
}

; 以当前全局参数生成方案对象
GetCurrentProfileObject() {
    global MarginType, MarginTop, MarginBottom, MarginLeft, MarginRight
    global CustomTop, CustomBottom, CustomLeft, CustomRight
    global PageNumFormat, SkipSinglePageNum
    global OptProtectTable, OptPreventOrphan, OptAttachmentNewPage, MaxHeadingLen, ActiveStyle
    global GridMode, GridLines, GridChars
    global HeaderMode, HeaderText, HeaderAlign, HeaderMarginTop, HeaderFont, HeaderSize, HeaderBold
    global HeaderUnderline, HeaderKeep
    obj := {
        MarginType: MarginType,
        MarginTop: MarginTop, MarginBottom: MarginBottom,
        MarginLeft: MarginLeft, MarginRight: MarginRight,
        CustomTop: CustomTop, CustomBottom: CustomBottom,
        CustomLeft: CustomLeft, CustomRight: CustomRight,
        PageNumFormat: PageNumFormat,
        SkipSinglePageNum: SkipSinglePageNum,
        GridMode: GridMode, GridLines: GridLines, GridChars: GridChars,
        CurrentFontScheme: GetBuiltinProfileName(),
        OptProtectTable: OptProtectTable,
        OptPreventOrphan: OptPreventOrphan,
        OptAttachmentNewPage: OptAttachmentNewPage,
        MaxHeadingLen: MaxHeadingLen,
        ; v6.8.20 页眉参数（v6.8.26：HeaderText 只有勾选“记住文字”才随方案保存）
        HeaderMode: HeaderMode,
        HeaderText: HeaderKeep ? HeaderText : "",
        HeaderAlign: HeaderAlign, HeaderMarginTop: HeaderMarginTop, HeaderFont: HeaderFont,
        HeaderSize: HeaderSize, HeaderBold: HeaderBold,
        HeaderUnderline: HeaderUnderline, HeaderKeep: HeaderKeep
    }
    ; 字体字号直接随方案保存
    for k in GetFontKeys()
        obj.%k% := ActiveStyle.%k%
    return obj
}

SaveProfileFromGlobals(name) {
    ProfileWriteObject(name, GetCurrentProfileObject())
}

; 把方案对象装载到运行时的全局参数
ApplyProfileToGlobals(obj) {
    global MarginType, MarginTop, MarginBottom, MarginLeft, MarginRight
    global CustomTop, CustomBottom, CustomLeft, CustomRight
    global PageNumFormat, SkipSinglePageNum, CurrentFontScheme
    global OptProtectTable, OptPreventOrphan, OptAttachmentNewPage, MaxHeadingLen
    global GridMode, GridLines, GridChars
    global HeaderMode, HeaderText, HeaderAlign, HeaderMarginTop, HeaderFont, HeaderSize, HeaderBold
    global HeaderUnderline, HeaderKeep
    MarginType := Integer(obj.MarginType)
    MarginTop := FmtMargin(obj.MarginTop)
    MarginBottom := FmtMargin(obj.MarginBottom)
    MarginLeft := FmtMargin(obj.MarginLeft)
    MarginRight := FmtMargin(obj.MarginRight)
    CustomTop := FmtMargin(obj.CustomTop)
    CustomBottom := FmtMargin(obj.CustomBottom)
    CustomLeft := FmtMargin(obj.CustomLeft)
    CustomRight := FmtMargin(obj.CustomRight)
    PageNumFormat := obj.PageNumFormat
    SkipSinglePageNum := Integer(obj.SkipSinglePageNum)
    CurrentFontScheme := GetBuiltinProfileName()
    OptProtectTable := Integer(obj.OptProtectTable)
    OptPreventOrphan := Integer(obj.OptPreventOrphan)
    OptAttachmentNewPage := Integer(obj.OptAttachmentNewPage)
    MaxHeadingLen := Integer(obj.MaxHeadingLen)
    GridMode := Integer(obj.GridMode)
    GridLines := Integer(obj.GridLines)
    GridChars := Integer(obj.GridChars)
    ; v6.8.20 页眉：空串是合法值（= 用文件名），所以这里不做“空则回落默认”的处理
    HeaderMode := Integer(obj.HeaderMode)
    HeaderKeep := Integer(obj.HeaderKeep)
    ; v6.8.26：未勾选“记住文字”的方案里存下来的旧页眉文字，装载时一律忽略；
    ; 这样即使注册表里残留着历史文字，也不会重新显示或被排版用到。
    HeaderText := HeaderKeep ? String(obj.HeaderText) : ""
    HeaderAlign := String(obj.HeaderAlign)
    HeaderMarginTop := FmtMargin(obj.HeaderMarginTop)
    HeaderFont := String(obj.HeaderFont)
    HeaderSize := String(obj.HeaderSize)
    HeaderBold := Integer(obj.HeaderBold)
    HeaderUnderline := Integer(obj.HeaderUnderline)

    ; 字体字号随方案一起装载
    ApplyStyleObject(obj)
}

; 把当前方案参数写到注册表顶层，供排版逻辑和静默模式读取
PersistProfileKeysToRegistry() {
    global MarginType, MarginTop, MarginBottom, MarginLeft, MarginRight
    global CustomTop, CustomBottom, CustomLeft, CustomRight
    global PageNumFormat, SkipSinglePageNum, CurrentFontScheme
    global OptProtectTable, OptPreventOrphan, OptAttachmentNewPage, MaxHeadingLen
    global GridMode, GridLines, GridChars
    global HeaderMode, HeaderText, HeaderAlign, HeaderMarginTop, HeaderFont, HeaderSize, HeaderBold
    global HeaderUnderline, HeaderKeep, REG_PATH
    SetConfigVal("Margins", "MarginType", MarginType)
    SetConfigVal("Margins", "MarginTop", FmtMargin(MarginTop))
    SetConfigVal("Margins", "MarginBottom", FmtMargin(MarginBottom))
    SetConfigVal("Margins", "MarginLeft", FmtMargin(MarginLeft))
    SetConfigVal("Margins", "MarginRight", FmtMargin(MarginRight))
    SetConfigVal("Margins", "CustomTop", FmtMargin(CustomTop))
    SetConfigVal("Margins", "CustomBottom", FmtMargin(CustomBottom))
    SetConfigVal("Margins", "CustomLeft", FmtMargin(CustomLeft))
    SetConfigVal("Margins", "CustomRight", FmtMargin(CustomRight))
    SetConfigVal("PageNumber", "PageNumFormat", PageNumFormat)
    SetConfigVal("PageNumber", "SkipSinglePageNum", SkipSinglePageNum)
    SetConfigVal("Margins", "GridMode", GridMode)
    SetConfigVal("Margins", "GridLines", GridLines)
    SetConfigVal("Margins", "GridChars", GridChars)
    SetConfigVal("FontScheme", "CurrentFontScheme", CurrentFontScheme)
    SetConfigVal("Options", "OptProtectTable", OptProtectTable)
    SetConfigVal("Options", "OptPreventOrphan", OptPreventOrphan)
    SetConfigVal("Options", "OptAttachmentNewPage", OptAttachmentNewPage)
    SetConfigVal("Options", "MaxHeadingLen", MaxHeadingLen)
    ; v6.8.20 页眉参数（顶层键；v6.8.26 起含 HeaderKeep），section 参数只是分组标签，实际都写在 REG_PATH 根下
    SetConfigVal("Header", "HeaderMode", HeaderMode)
    ; v6.8.26：只有勾选“记住文字”才保存 HeaderText；否则写回空串，
    ; 避免单位名、文件主题残留到下一次排版或右键批量。
    if (HeaderKeep)
        SetConfigVal("Header", "HeaderText", HeaderText)
    else
        SetConfigVal("Header", "HeaderText", "")
    SetConfigVal("Header", "HeaderKeep", HeaderKeep)
    SetConfigVal("Header", "HeaderAlign", HeaderAlign)
    SetConfigVal("Header", "HeaderMarginTop", FmtMargin(HeaderMarginTop))
    SetConfigVal("Header", "HeaderFont", HeaderFont)
    SetConfigVal("Header", "HeaderSize", HeaderSize)
    SetConfigVal("Header", "HeaderBold", HeaderBold)
    SetConfigVal("Header", "HeaderUnderline", HeaderUnderline)
}

ApplyProfile(name, showToast := true) {
    global ActiveProfile, uiSB
    if (name == "")
        name := GetBuiltinProfileName()

    obj := ProfileRead(name)
    ApplyProfileToGlobals(obj)
    ActiveProfile := name
    SetConfigVal("Profile", "Profile_Active", ActiveProfile)
    PersistProfileKeysToRegistry()

    if (IsSet(mainGui) && IsObject(mainGui)) {
        RefreshProfileDropdown()
        RefreshProfileInfo()
        UpdateProfileButtons()
        if (IsSet(uiSB) && IsObject(uiSB))
            uiSB.SetText(" 当前方案：" ActiveProfile)
        if (showToast)
            ShowFloatingToast("已切换方案", ActiveProfile, 1800, true)
    }
}

; 启动时校验当前方案：没有用户方案或记录失效时，回到内置“党政机关公文格式”
EnsureProfiles() {
    global ActiveProfile
    names := GetUserProfileNames()
    active := GetConfigVal("Profile", "Profile_Active", "")

    if (active == "" || (!IsBuiltinProfile(active) && !ProfileExists(active))) {
        if (names.Length > 0)
            active := names[1]
        else
            active := GetBuiltinProfileName()
    }

    ActiveProfile := active
    ApplyProfile(active, false)
}

GetMarginTypeLabel(type) {
    switch Integer(type) {
        case 2: return "WORD 默认边距"
        case 3: return "用户自定义"
        default: return "党政机关公文格式"
    }
}

; 首页“页眉处理”摘要文字（v6.8.20；v6.8.22 追加页眉横线状态）
; 只读全局参数，不改任何状态；页眉文字过长时截断，避免把首页那一行撑出组框。
GetHeaderSummaryText() {
    global HeaderMode, HeaderText, HeaderAlign, HeaderMarginTop, HeaderUnderline
    switch Integer(HeaderMode) {
        case 1:
            txt := Trim(HeaderText)
            if (txt == "")
                txt := "用文件名"
            else if (StrLen(txt) > 16)
                txt := SubStr(txt, 1, 16) "…"
            ; 页眉横线只在“添加·替换”下有实际意义（另外两种不写页眉格式）
            ulText := (Integer(HeaderUnderline) != 0) ? " · 页眉横线" : ""
            return "添加·替换 · " txt " · " HeaderAlign " · 距上边 " FmtMargin(HeaderMarginTop) " mm" ulText
        case 2:
            return "清空（删除原有页眉）"
        default:
            return "不改变（保留原有页眉）"
    }
}

; 刷新首页方案下拉列表（内置方案始终第一项）
RefreshProfileDropdown() {
    global uiProfile, ActiveProfile, isProgramUpdating
    if (!IsSet(uiProfile) || !IsObject(uiProfile))
        return

    names := [GetBuiltinProfileName()]
    for n in GetUserProfileNames()
        names.Push(n)

    if (ActiveProfile == "" || (!IsBuiltinProfile(ActiveProfile) && !ProfileExists(ActiveProfile)))
        ActiveProfile := names.Length > 1 ? names[2] : names[1]

    isProgramUpdating := true
    try uiProfile.Delete()
    uiProfile.Add(names)
    targetIdx := 1
    for idx, n in names {
        if (n == ActiveProfile) {
            targetIdx := idx
            break
        }
    }
    uiProfile.Choose(targetIdx)
    isProgramUpdating := false
    UpdateProfileButtons()
}

; 刷新首页“当前方案信息”只读摘要
RefreshProfileInfo() {
    global uiInfoMarginType, uiInfoMargins, uiInfoPage, uiInfoSkip, uiInfoFont, uiInfoLen, uiInfoOptions, uiInfoHeader
    global MarginType, MarginTop, MarginBottom, MarginLeft, MarginRight
    global PageNumFormat, SkipSinglePageNum, ActiveStyle
    global OptProtectTable, OptPreventOrphan, OptAttachmentNewPage, MaxHeadingLen
    global HeaderMode, HeaderText, HeaderAlign, HeaderMarginTop
    if (!IsSet(uiInfoMarginType) || !IsObject(uiInfoMarginType))
        return

    uiInfoMarginType.Text := GetMarginTypeLabel(MarginType)
    uiInfoMargins.Text := MarginTop "/" MarginBottom "/" MarginLeft "/" MarginRight " mm"
    uiInfoPage.Text := PageNumFormat
    uiInfoSkip.Text := SkipSinglePageNum ? "是" : "否"
    uiInfoFont.Text := ActiveStyle.BodyFont " " ActiveStyle.BodySize "，行距 " ActiveStyle.Spacing
    uiInfoLen.Text := MaxHeadingLen " 字"
    uiInfoOptions.Text := "表格保护 " (OptProtectTable ? "✓" : "✗")
        . "     防孤行 " (OptPreventOrphan ? "✓" : "✗")
        . "     附件另面 " (OptAttachmentNewPage ? "✓" : "✗")
    ; v6.8.20 新增第五行：页眉处理方式摘要（本版只显示，不参与排版）
    uiInfoHeader.Text := GetHeaderSummaryText()
    uiInfoHeader.SetFont("s10 Bold c000000")   ; 页眉处理方式信息加粗，更显著
}

; 内置只读方案不可编辑；“另存为”始终可用
UpdateProfileButtons() {
    global uiBtnEditProfile, uiBtnSaveAsProfile, ActiveProfile
    if (!IsSet(uiBtnEditProfile) || !IsObject(uiBtnEditProfile))
        return
    isBuiltin := IsBuiltinProfile(ActiveProfile)
    uiBtnEditProfile.Enabled := !isBuiltin
    uiBtnSaveAsProfile.Enabled := true
    try uiBtnEditProfile.ToolTip := isBuiltin ? "内置只读方案，请使用“另存为…”创建自己的方案" : "编辑当前方案的参数"
}

OnProfileChange(ctrl, *) {
    global isProgramUpdating
    if (isProgramUpdating)
        return
    name := ctrl.Text
    if (name == "")
        return
    ApplyProfile(name)
}

OnBackupCheckboxChange(ctrl, *) {
    global CreateBAK
    CreateBAK := ctrl.Value ? 1 : 0
    SetConfigVal("Options", "CreateBAK", CreateBAK)
}

ValidateProfileName(name) {
    if (name == "") {
        MsgBox("方案名称不能为空。", "提示", "Icon!")
        return false
    }
    if (InStr(name, "\")) {
        MsgBox("方案名称不能包含反斜杠“\”。", "提示", "Icon!")
        return false
    }
    if (StrLen(name) > 60) {
        MsgBox("方案名称过长，请控制在 60 个字符以内。", "提示", "Icon!")
        return false
    }
    return true
}

SaveProfileAs(*) {
    global ActiveProfile
    wasBuiltin := IsBuiltinProfile(ActiveProfile)
    defaultName := wasBuiltin ? "我的方案" : ActiveProfile " 副本"

    ib := InputBox("请输入新方案的名称（例如：XX局 · 通知）", "另存为方案", "w420 h160", defaultName)
    if (ib.Result != "OK")
        return

    newName := Trim(ib.Value)
    if (!ValidateProfileName(newName))
        return
    if (IsBuiltinProfile(newName)) {
        MsgBox("“" GetBuiltinProfileName() "”为内置只读方案，请换一个名称。", "提示", "Icon!")
        return
    }
    if (ProfileExists(newName)) {
        if (MsgBox("方案“" newName "”已存在，是否覆盖？", "确认覆盖", "YesNo Icon?") != "Yes")
            return
    }

    SaveProfileFromGlobals(newName)
    ApplyProfile(newName)
    ShowFloatingToast("方案已保存", "已另存为：" newName, 2000, true)

    ; 从内置只读方案另存后，直接打开编辑窗口方便继续调整
    if (wasBuiltin)
        OpenProfileEditor()
}

; =====================================================================
; [国标对照：字体方案管理窗口中给偏离国标的行加星号]
; =====================================================================
NormalizeFontNameForCompare(name) {
    n := Trim(name)
    for suffix in ["_GB2312", "_GBK", " GB2312", " GBK"] {
        if (StrLen(n) > StrLen(suffix) && SubStr(n, -StrLen(suffix) + 1) == suffix)
            n := SubStr(n, 1, StrLen(n) - StrLen(suffix))
    }
    return n
}

; 国标正文行距为“固定值 28 磅”；带“倍”的不是国标写法
IsStandardSpacingText(text) {
    t := StrReplace(StrReplace(Trim(text), " ", ""), "　", "")
    if (t == "" || InStr(t, "倍"))
        return false
    if RegExMatch(t, "(\d+(\.\d+)?)", &m)
        return Abs(Float(m[1]) - 28) < 0.01
    return false
}

; 逐行对照国标字体、字号、加粗，以及正文行距，偏离时在行名称后显示红色星号
UpdateFontSchemeDeviation(*) {
    global fontMgrRows, fontMgrSpacingCtrl, fontMgrSpacingStar
    if (IsSet(fontMgrRows) && IsObject(fontMgrRows)) {
        for row in fontMgrRows {
            fontOk := NormalizeFontNameForCompare(row.font.Text) == NormalizeFontNameForCompare(row.stdFont)
            sizeOk := Trim(row.size.Text) == row.stdSize
            boldOk := Integer(row.bold.Value) == Integer(row.stdBold)
            SetDeviationStar(row.star, !(fontOk && sizeOk && boldOk))
        }
    }
    if (IsSet(fontMgrSpacingCtrl) && IsObject(fontMgrSpacingCtrl))
        SetDeviationStar(fontMgrSpacingStar, !IsStandardSpacingText(fontMgrSpacingCtrl.Text))
}

; =====================================================================
; [自定义字体字号方案管理子窗口]
; =====================================================================
OpenFontSchemeManager(*) {
    global mainGui, editGui, CustomFontSchemes, dlgFontScheme, CurrentFontScheme
    global WpsFontNameList, WpsFontSizeList, LineSpacingPresets
    global fontMgrRows, fontMgrSpacingCtrl, fontMgrSpacingLabel, fontMgrSpacingStar

    ; 每次打开窗口时重建“行名称 → 控件”对照表，用于把偏离国标的行标星
    fontMgrRows := []
    fontMgrSpacingCtrl := "", fontMgrSpacingLabel := "", fontMgrSpacingStar := ""

    ownerGui := (IsSet(editGui) && IsObject(editGui)) ? editGui : mainGui
    curScheme := (IsSet(dlgFontScheme) && IsObject(dlgFontScheme)) ? dlgFontScheme.Text : CurrentFontScheme

    mGui := Gui("+Owner" ownerGui.Hwnd " +ToolWindow", "自定义字体字号方案管理")
    mGui.SetFont("s9", "Microsoft YaHei")

    mGui.Add("Text", "x20 y15 w70", "已有方案:")
    schemeItems := []
    for sName in CustomFontSchemes
        schemeItems.Push(sName)
    lbSchemes := mGui.Add("ListBox", "x20 y35 w150 h415", schemeItems)

    mGui.Add("GroupBox", "x185 y15 w520 h435", "方案参数配置 (各级文本对象)")
    mGui.Add("Text", "x200 y40 w65", "方案名称:")
    edName := mGui.Add("Edit", "x270 y37 w160")

    itemsDef := [
        ["主标题:",   "MTitle",   "方正小标宋简体", "二号", 0, "小标宋 二号 不加粗"],
        ["副标题:",   "SubTitle", "楷体",           "三号", 0, "楷体 三号 不加粗"],
        ["作者/发文:", "Author",   "楷体",           "三号", 0, "楷体 三号 不加粗"],
        ["正文内容:", "Body",     "仿宋",           "三号", 0, "仿宋 三号 不加粗"],
        ["一级标题:", "H1",       "黑体",           "三号", 0, "黑体 三号 不加粗"],
        ["二级标题:", "H2",       "楷体",           "三号", 0, "楷体 三号 不加粗"],
        ["三级标题:", "H3",       "仿宋",           "三号", 1, "仿宋 三号 (常加粗)"],
        ["四级标题:", "H4",       "仿宋",           "三号", 0, "仿宋 三号 不加粗"],
        ["表格内容:", "Table",    "仿宋",           "五号", 0, "仿宋 五号 不加粗"]
    ]

    ctrlMap := Map()
    curY := 70

    for it in itemsDef {
        lblRow := mGui.Add("Text", "x200 y" curY + 3 " w65", it[1])
        starRow := AddDeviationStar(mGui, "x260 y" curY + 3 " w10")
        cbF := mGui.Add("ComboBox", "x270 y" curY " w120", WpsFontNameList)
        cbF.Text := it[3]
        cbS := mGui.Add("ComboBox", "x395 y" curY " w75", WpsFontSizeList)
        cbS.Text := it[4]
        chkB := mGui.Add("Checkbox", "x478 y" curY + 3 " w50 Checked" it[5], "加粗")

        mGui.SetFont("s8 c777777")
        mGui.Add("Text", "x535 y" curY + 3 " w160", "（国标: " it[6] "）")
        mGui.SetFont("s9 c000000")

        ctrlMap[it[2] "Font"] := cbF
        ctrlMap[it[2] "Size"] := cbS
        ctrlMap[it[2] "Bold"] := chkB

        ; 记录每行控件与国标值，任一项偏离时在名称后显示红色星号
        fontMgrRows.Push({label: lblRow, star: starRow, font: cbF, size: cbS, bold: chkB,
                          stdFont: it[3], stdSize: it[4], stdBold: it[5]})
        cbF.OnEvent("Change", UpdateFontSchemeDeviation)
        cbS.OnEvent("Change", UpdateFontSchemeDeviation)
        chkB.OnEvent("Click", UpdateFontSchemeDeviation)
        curY += 33
    }

    fontMgrSpacingLabel := mGui.Add("Text", "x200 y" curY + 3 " w65", "正文行距:")
    fontMgrSpacingStar := AddDeviationStar(mGui, "x260 y" curY + 3 " w10")
    cbSpacing := mGui.Add("ComboBox", "x270 y" curY " w160", LineSpacingPresets)
    cbSpacing.Text := "固定值 28 磅"
    mGui.SetFont("s8 c777777")
    mGui.Add("Text", "x440 y" curY + 3 " w250", "（未输入'倍'默认按磅，国标: 28 磅）")
    mGui.SetFont("s9 c000000")
    fontMgrSpacingCtrl := cbSpacing
    cbSpacing.OnEvent("Change", UpdateFontSchemeDeviation)

    LoadSelectedSchemeToEditor() {
        sel := lbSchemes.Text
        if (sel != "" && CustomFontSchemes.Has(sel)) {
            data := CustomFontSchemes[sel]
            edName.Value := sel
            for k, ctrl in ctrlMap {
                if InStr(k, "Bold")
                    ctrl.Value := Integer(data.%k%)
                else
                    ctrl.Text := data.%k%
            }
            cbSpacing.Text := data.Spacing
            UpdateFontSchemeDeviation()
        }
    }

    lbSchemes.OnEvent("Change", (*) => LoadSelectedSchemeToEditor())

    if (curScheme != "党政机关公文格式" && CustomFontSchemes.Has(curScheme)) {
        for idx, sName in schemeItems {
            if (sName == curScheme) {
                lbSchemes.Choose(idx)
                LoadSelectedSchemeToEditor()
                break
            }
        }
    }

    btnSave := mGui.Add("Button", "x280 y460 w160 h32 Default", "保存/新建此方案")
    btnDelete := mGui.Add("Button", "x470 y460 w150 h32", "删除选中方案")

    OnSaveSchemeClick(*) {
        global fontMgrRows, fontMgrSpacingCtrl, fontMgrSpacingLabel, fontMgrSpacingStar
        sName := Trim(edName.Value)
        if (sName == "" || sName == "党政机关公文格式") {
            MsgBox("请输入有效的方案名称，且不能与默认格式重名！", "提示", "Icon!")
            return
        }
        packObj := {Spacing: Trim(cbSpacing.Text)}
        for k, ctrl in ctrlMap
            packObj.%k% := InStr(k, "Bold") ? Integer(ctrl.Value) : Trim(ctrl.Text)

        SaveFontSchemeData(sName, packObj)
        CustomFontSchemes[sName] := packObj
        RefreshFontSchemeDropdown(sName)
        fontMgrRows := [], fontMgrSpacingCtrl := "", fontMgrSpacingLabel := "", fontMgrSpacingStar := ""
        mGui.Destroy()
    }

    OnDeleteSchemeClick(*) {
        global fontMgrRows, fontMgrSpacingCtrl, fontMgrSpacingLabel, fontMgrSpacingStar
        sel := lbSchemes.Text
        if (sel != "" && CustomFontSchemes.Has(sel)) {
            DeleteFontSchemeData(sel)
            CustomFontSchemes.Delete(sel)
            RefreshFontSchemeDropdown("党政机关公文格式")
            fontMgrRows := [], fontMgrSpacingCtrl := "", fontMgrSpacingLabel := "", fontMgrSpacingStar := ""
            mGui.Destroy()
        }
    }

    btnSave.OnEvent("Click", OnSaveSchemeClick)
    btnDelete.OnEvent("Click", OnDeleteSchemeClick)
    UpdateFontSchemeDeviation()
    mGui.Show("w720 h510")
}

RefreshFontSchemeDropdown(activeScheme := "党政机关公文格式") {
    global dlgFontScheme, CustomFontSchemes
    ; 只刷新方案编辑窗口中的字体下拉，不写注册表、不立即应用；
    ; 等用户点“保存”后再随方案一起生效。
    if (!IsSet(dlgFontScheme) || !IsObject(dlgFontScheme))
        return

    newList := ["党政机关公文格式"]
    for sName in CustomFontSchemes
        newList.Push(sName)
    dlgFontScheme.Delete()
    dlgFontScheme.Add(newList)

    targetIdx := 1
    for idx, item in newList {
        if (item == activeScheme) {
            targetIdx := idx
            break
        }
    }
    dlgFontScheme.Choose(targetIdx)
    UpdateProfileDeviationHighlight()
}

; =====================================================================
; [整体配置方案：编辑参数窗口]
; =====================================================================
FmtMargin(v) {
    n := IsNumber(v) ? Float(v) : 0.0
    return (Round(n, 1) == Round(n)) ? String(Round(n)) : Format("{:.1f}", n)
}

; 国标（GB/T 9704-2012）关键参数，用于编辑窗口中判断偏差项
GetStandardMargin(which) {
    switch which {
        case "Top": return 37
        case "Bottom": return 35
        case "Left": return 28
        case "Right": return 26
    }
    return 0
}

; 与国标不一致时，在选项名称后显示红色星号（名称本身保持黑色）
SetDeviationStar(starCtrl, isDeviating) {
    if (!IsSet(starCtrl) || !IsObject(starCtrl))
        return
    starCtrl.Visible := !!isDeviating
}

; 创建一个默认隐藏的红色星号，紧跟名称标签显示
AddDeviationStar(guiObj, options) {
    star := guiObj.Add("Text", options, "*")
    star.SetFont("s9 cFF0000", "Microsoft YaHei")
    star.Visible := false
    return star
}

; ---------------------------------------------------------------------
; 创建一个与上一控件（Edit）自动绑定的微调按钮（UpDown）
;
; 参数：guiObj  —— 目标窗口
;       editCtrl —— 已经创建好、需要绑定的 Edit 控件对象
;       rangeOpt —— 取值范围，如 "Range1-100"
;       value    —— 输入框当前值（数字，可能是小数）
;
; v6.8.19 修复的三个坑（都是实测出来的，改之前务必看完）：
;
;   1. 选项里绝对不能再写 "-2"。
;      AHK v2 会把选项串里的裸数字当成“窗口样式位”：正数加上、负数减掉，
;      -2 正好减掉 UDS_SETBUDDYINT(0x0002)——这个位就是“把 UpDown 的数值
;      写回 buddy 输入框”的开关。老代码六个微调按钮全带 -2，于是箭头点击
;      只改 UpDown 自己的内部数值，Edit 文本一个字符都不变，用户看到的就是
;      “点了没反应”。实测样式：带 -2 是 0x50000034，去掉是 0x50000036。
;
;   2. 不写 x/y 坐标时，AHK 会把“上一个添加的控件”自动设为 buddy，并把
;      Edit 右侧约 17 像素让给箭头（所以 Edit 实际可见宽度比写的 w 少 17）。
;      这也是本函数要求“先建 Edit、紧接着建 UpDown”的原因。
;
;   3. 创建 UpDown 时传的初始值会立刻被写进 Edit 文本（带 SETBUDDYINT 时）。
;      所以必须传当前值，不能像老代码那样硬写 50，否则窗口一打开四个边距
;      就全变成 50。又因为 UpDown 内部只存整数，而 Word 默认边距是 25.4 这种
;      小数，这里先传四舍五入后的整数让箭头有个合理起点，随后立刻把 Edit
;      文本恢复成精确原值，避免 25.4 被显示成 25（那会被误判成偏离国标）。
; ---------------------------------------------------------------------
MakeSpin(guiObj, editCtrl, rangeOpt, value) {
    num := IsNumber(value) ? Float(value) : 0
    ud := guiObj.Add("UpDown", rangeOpt, Round(num))
    try editCtrl.Value := value      ; 恢复原始文本（小数不被截断）
    return ud
}

; 逐个检查编辑窗口中的边距、字体字号是否偏离国标，并显示对应星号
UpdateProfileDeviationHighlight(*) {
    global dlgTop, dlgBottom, dlgLeft, dlgRight, dlgMarginType
    global dlgStarTop, dlgStarBottom, dlgStarLeft, dlgStarRight, dlgStarMarginType
    global dlgFontRows, dlgSpacingCtrl, dlgSpacingStar
    if (!IsSet(dlgTop) || !IsObject(dlgTop) || !IsObject(dlgMarginType))
        return

    for pair in [["Top", dlgTop, dlgStarTop], ["Bottom", dlgBottom, dlgStarBottom],
                 ["Left", dlgLeft, dlgStarLeft], ["Right", dlgRight, dlgStarRight]] {
        text := Trim(pair[2].Value)
        std := GetStandardMargin(pair[1])
        deviates := !IsNumber(text) || Abs(Float(text) - std) > 0.05
        SetDeviationStar(pair[3], deviates)
    }

    ; 边距预设本身不是国标预设时显示星号
    SetDeviationStar(dlgStarMarginType, dlgMarginType.Text != "党政机关公文格式")

    ; 字体字号逐行对照国标，偏离时在名称后显示星号
    if (IsSet(dlgFontRows) && IsObject(dlgFontRows)) {
        for row in dlgFontRows {
            fontOk := NormalizeFontNameForCompare(row.font.Text) == NormalizeFontNameForCompare(row.stdFont)
            sizeOk := Trim(row.size.Text) == row.stdSize
            boldOk := Integer(row.bold.Value) == Integer(row.stdBold)
            SetDeviationStar(row.star, !(fontOk && sizeOk && boldOk))
        }
    }
    if (IsSet(dlgSpacingCtrl) && IsObject(dlgSpacingCtrl))
        SetDeviationStar(dlgSpacingStar, !IsStandardSpacingText(dlgSpacingCtrl.Text))
}


OpenProfileEditor(*) {
    global mainGui, editGui, editorOriginalName, ActiveProfile
    global dlgName, dlgMarginType, dlgTop, dlgBottom, dlgLeft, dlgRight
    global dlgLblMarginType, dlgLblTop, dlgLblBottom, dlgLblLeft, dlgLblRight
    global dlgStarMarginType, dlgStarTop, dlgStarBottom, dlgStarLeft, dlgStarRight
    global dlgTopUD, dlgBottomUD, dlgLeftUD, dlgRightUD
    global dlgPageFormat, dlgSkipSingle
    global dlgProtectTable, dlgPreventOrphan, dlgAttachNewPage, dlgMaxHeadingLen
    global dlgCustomTop, dlgCustomBottom, dlgCustomLeft, dlgCustomRight
    global dlgFontRows, dlgSpacingCtrl, dlgSpacingStar, dlgSpacingLabel
    global dlgGridMode, dlgGridLines, dlgGridChars, dlgGridLinesUD, dlgGridCharsUD
    global PageFormatList, ActiveStyle
    global MarginType, MarginTop, MarginBottom, MarginLeft, MarginRight
    global CustomTop, CustomBottom, CustomLeft, CustomRight
    global PageNumFormat, SkipSinglePageNum
    global GridMode, GridLines, GridChars
    global OptProtectTable, OptPreventOrphan, OptAttachmentNewPage, MaxHeadingLen
    ; v6.8.20 页眉：控件对象 + 当前生效值（v6.8.26：含记住文字复选框）
    global dlgHeaderMode, dlgHeaderText, dlgHeaderAlign
    global dlgHeaderFont, dlgHeaderSize, dlgHeaderBold, dlgHeaderUnderline
    global dlgHeaderMargin, dlgHeaderMarginUD, dlgHeaderKeep
    global HeaderMode, HeaderText, HeaderAlign, HeaderMarginTop, HeaderFont, HeaderSize, HeaderBold
    global HeaderUnderline, HeaderKeep

    if (IsBuiltinProfile(ActiveProfile)) {
        MsgBox("“" GetBuiltinProfileName() "”是内置只读方案。`n`n请使用“另存为…”创建可编辑的方案。", "只读方案", "Icon!")
        return
    }
    if (IsSet(editGui) && IsObject(editGui)) {
        editGui.Show()
        try WinActivate("ahk_id " editGui.Hwnd)
        return
    }

    editorOriginalName := ActiveProfile
    ; 编辑器内的“自定义边距”缓存：取消时不写回运行时参数
    dlgCustomTop := CustomTop, dlgCustomBottom := CustomBottom
    dlgCustomLeft := CustomLeft, dlgCustomRight := CustomRight

    editGui := Gui("+Owner" mainGui.Hwnd " +ToolWindow", "编辑方案参数 — " ActiveProfile)
    editGui.SetFont("s9", "Microsoft YaHei")
    editGui.BackColor := "FFFFFF"
    editGui.OnEvent("Close", (*) => CloseProfileEditor())
    editGui.OnEvent("Escape", (*) => CloseProfileEditor())

    editGui.Add("Text", "x20 y19 w70", "方案名称:")
    dlgName := editGui.Add("Edit", "x90 y15 w290", ActiveProfile)

    ; --- 页面设置：页边距 + 文档网格（左列，v6.8.18 合并 / v6.8.19 重排）---
    ; v6.8.19 统一栅格：左右外边距 20、两列间距 20（左列 20–380、右列 400–900），
    ; 组框内控件统一左右各留 12（32 → 368）。
    ; 边距区是真正的等宽两列：列 1 起于 x32、列 2 起于 x200，每列内部固定为
    ; 「标签 56 + 星号 10 + 输入框 78 + 单位 24」，两列的输入框与单位列上下严格对齐，
    ; 不再出现左边留 12、右边留 26 的视觉偏移。
    ; 组框内分两段：上半页边距、下半文档网格，中间一条 1 像素浅灰线分隔；
    ; 两段各有一个灰色小节标题，与右侧“字体字号”表头同款。
    ; 组框内两段各自留出 30/32 像素的行距（6.8.19 第一次预览后按反馈加大），
    ; 组框高度 258（y50–y308），与下方两个组框、右侧字体组框一起把底边收在 y564。
    editGui.Add("GroupBox", "x20 y50 w360 h258", "页面设置")
    editGui.SetFont("s8 c666666")
    editGui.Add("Text", "x32 y70 w150", "页边距 (毫米)")
    editGui.SetFont("s9 c000000")
    dlgLblMarginType := editGui.Add("Text", "x32 y98 w56", "边距预设:")
    dlgStarMarginType := AddDeviationStar(editGui, "x88 y98 w10")
    dlgMarginType := editGui.Add("DropDownList", "x98 y94 w150 Choose" ((MarginType >= 1 && MarginType <= 3) ? MarginType : 1), ["党政机关公文格式", "WORD默认边距", "用户自定义"])
    dlgMarginType.OnEvent("Change", OnDialogMarginTypeChange)

    ; 上/下边距一行、左/右边距一行：两行的列位置完全一致，只是 y 相差 26，
    ; 保证四组“标签 + 星号 + 数字框 + mm”在视觉上落成规整的 2×2 网格。
    dlgLblTop := editGui.Add("Text", "x32 y128 w56", "上边距:")
    dlgStarTop := AddDeviationStar(editGui, "x88 y128 w10")
    dlgTop := editGui.Add("Edit", "x98 y124 w78 Center", MarginTop)
    dlgTopUD := MakeSpin(editGui, dlgTop, "Range0-100", MarginTop)
    editGui.Add("Text", "x176 y128 w24", "mm")
    dlgLblBottom := editGui.Add("Text", "x200 y128 w56", "下边距:")
    dlgStarBottom := AddDeviationStar(editGui, "x256 y128 w10")
    dlgBottom := editGui.Add("Edit", "x266 y124 w78 Center", MarginBottom)
    dlgBottomUD := MakeSpin(editGui, dlgBottom, "Range0-100", MarginBottom)
    editGui.Add("Text", "x344 y128 w24", "mm")

    dlgLblLeft := editGui.Add("Text", "x32 y158 w56", "左边距:")
    dlgStarLeft := AddDeviationStar(editGui, "x88 y158 w10")
    dlgLeft := editGui.Add("Edit", "x98 y154 w78 Center", MarginLeft)
    dlgLeftUD := MakeSpin(editGui, dlgLeft, "Range0-100", MarginLeft)
    editGui.Add("Text", "x176 y158 w24", "mm")
    dlgLblRight := editGui.Add("Text", "x200 y158 w56", "右边距:")
    dlgStarRight := AddDeviationStar(editGui, "x256 y158 w10")
    dlgRight := editGui.Add("Edit", "x266 y154 w78 Center", MarginRight)
    dlgRightUD := MakeSpin(editGui, dlgRight, "Range0-100", MarginRight)
    editGui.Add("Text", "x344 y158 w24", "mm")

    ; 页边距与文档网格之间的细分隔线：宽度撑满组框内宽（32–368）。
    ; 用 1 像素高的浅灰静态文本模拟分隔线，不可交互，也不参与 Tab 顺序。
    editGui.Add("Text", "x32 y186 w336 h1 BackgroundDDDDDD")

    ; --- 文档网格（仍属“页面设置”组，放在页边距下方）---
    ; 网格类型独占一行（下拉框 170 宽）；行数、字符数并排，单位列与上方“mm”同列，
    ; 这样“每页行数/每行字符数”读起来和边距四格是同一套对齐规则。
    editGui.SetFont("s8 c666666")
    editGui.Add("Text", "x32 y194 w150", "文档网格")
    editGui.SetFont("s9 c000000")
    editGui.Add("Text", "x32 y222 w64", "网格类型:")
    gridModeIdx := (Integer(GridMode) == 1) ? 1 : ((Integer(GridMode) == 2) ? 2 : 3)
    dlgGridMode := editGui.Add("DropDownList", "x98 y218 w170 Choose" gridModeIdx,
        ["指定行和字符网格", "只指定行网格", "不指定网格"])
    dlgGridMode.OnEvent("Change", OnDialogGridChange)

    editGui.Add("Text", "x32 y252 w64", "每页行数:")
    dlgGridLines := editGui.Add("Edit", "x98 y248 w78 Center Number", GridLines)
    dlgGridLinesUD := MakeSpin(editGui, dlgGridLines, "Range1-100", GridLines)
    editGui.Add("Text", "x176 y252 w24", "行")

    editGui.Add("Text", "x200 y252 w70", "每行字符数:")
    dlgGridChars := editGui.Add("Edit", "x272 y248 w70 Center Number", GridChars)
    dlgGridCharsUD := MakeSpin(editGui, dlgGridChars, "Range0-100", GridChars)
    editGui.Add("Text", "x344 y252 w20", "个")

    editGui.SetFont("s8 c999999")
    editGui.Add("Text", "x32 y280 w336", "国标：1 / 22 / 28　　Word 默认：2 / 44 / 不指定")
    editGui.SetFont("s9 c000000")

    ; --- 页码设置（左列）---
    ; v6.8.19 组框下移到 y320、高 86；页码下拉框加宽到 270，右端正好顶到内边界 368。
    editGui.Add("GroupBox", "x20 y320 w360 h86", "页码设置")
    editGui.Add("Text", "x32 y346 w60", "页码格式:")
    chooseFormatIdx := 6
    for idx, item in PageFormatList {
        if (item == PageNumFormat) {
            chooseFormatIdx := idx
            break
        }
    }
    dlgPageFormat := editGui.Add("DropDownList", "x98 y342 w270 Choose" chooseFormatIdx, PageFormatList)
    dlgSkipSingle := editGui.Add("Checkbox", "x32 y373 w336 Checked" SkipSinglePageNum, "文件仅一页不添加页码")

    ; --- 排版选项（左列）---
    ; v6.8.18 删除“规范发文字号六角括号”后只剩 3 个复选框，组框改名“排版选项”；
    ; v6.8.19 组框下移到 y418、高 146，三个复选框行距统一 25，
    ; 底边 564 与右侧字体组框对齐，左列不再比右列长出一截。
    editGui.Add("GroupBox", "x20 y418 w360 h146", "排版选项")
    dlgProtectTable := editGui.Add("Checkbox", "x32 y441 w336 Checked" OptProtectTable, "表格保护与字号适配")
    dlgPreventOrphan := editGui.Add("Checkbox", "x32 y466 w336 Checked" OptPreventOrphan, "标题与下文同页 (防孤行)")
    dlgAttachNewPage := editGui.Add("Checkbox", "x32 y491 w336 Checked" OptAttachmentNewPage, "附件另起一页 (附件另面编排)")

    editGui.Add("Text", "x32 y524 w60", "标题阈值:")
    dlgMaxHeadingLen := editGui.Add("Edit", "x98 y520 w78 Center Number", MaxHeadingLen)
    MakeSpin(editGui, dlgMaxHeadingLen, "Range10-200", MaxHeadingLen)
    editGui.Add("Text", "x180 y524 w188", "字（超过此长度按正文处理）")

    ; --- 字体字号（右列，随方案一起保存）---
    ; v6.8.19：右列整体右移 5 像素（x400 w500），与左列间距统一为 20；
    ; 五列位置固定：项目 415 / 字体 486 / 字号 612 / 加粗 676 / 国标参考 726，
    ; 表头与数据行同列，右端留白 16。行距 30→36 让 11 行不再拥挤，
    ; 同时把组框加高到 514，与左列一起收在 y564。
    ; “附件标记”“正文行距”前各有一条细分隔线，把三段内容在视觉上分开。
    editGui.Add("GroupBox", "x400 y50 w500 h514", "字体字号（随方案一起保存）")
    editGui.SetFont("s8 c666666")
    editGui.Add("Text", "x415 y64 w58", "项目")
    editGui.Add("Text", "x486 y64 w120", "字体")
    editGui.Add("Text", "x612 y64 w58", "字号")
    editGui.Add("Text", "x676 y64 w48", "加粗")
    editGui.Add("Text", "x726 y64 w158", "国标参考")
    editGui.SetFont("s9 c000000")

    dlgFontRows := []
    curY := 86
    for it in GetFontItemsDef() {
        ; 在“附件标记”行之前画一条 1 像素浅灰细分隔线，
        ; 把正文各级标题与“附件标记/页码”两段视觉上分开。
        ; 分隔线本身占 22 像素（线 + 上下留白），下面的 curY 增加量要与之匹配，
        ; 否则后续行会压到分隔线上。
        if (it[2] == "AttachMark") {
            editGui.Add("Text", "x415 y" (curY + 6) " w470 h1 BackgroundDDDDDD")
            curY += 22
        }
        prefix := it[2]
        fontKey := prefix . "Font"
        sizeKey := prefix . "Size"
        boldKey := prefix . "Bold"
        lblRow := editGui.Add("Text", "x415 y" curY + 4 " w58", it[1])
        starRow := AddDeviationStar(editGui, "x473 y" curY + 4 " w10")
        cbF := editGui.Add("ComboBox", "x486 y" curY " w120", WpsFontNameList)
        cbF.Text := ActiveStyle.%fontKey%
        cbS := editGui.Add("ComboBox", "x612 y" curY " w58", WpsFontSizeList)
        cbS.Text := ActiveStyle.%sizeKey%
        chkB := editGui.Add("Checkbox", "x676 y" curY + 4 " w48 h18 Checked" ActiveStyle.%boldKey%, "加粗")

        editGui.SetFont("s8 c999999")
        ; 去掉外层“（国标: …）”括号，只留“国标 …”，视觉更干净
        editGui.Add("Text", "x726 y" curY + 5 " w158", "国标 " it[6])
        editGui.SetFont("s9 c000000")

        dlgFontRows.Push({label: lblRow, star: starRow, font: cbF, size: cbS, bold: chkB,
                          prefix: prefix, stdFont: it[3], stdSize: it[4], stdBold: it[5]})
        cbF.OnEvent("Change", UpdateProfileDeviationHighlight)
        cbS.OnEvent("Change", UpdateProfileDeviationHighlight)
        chkB.OnEvent("Click", UpdateProfileDeviationHighlight)
        curY += 36
    }

    ; 行距行按 curY 动态定位：字体行数量会随版本增加（v6.8.13 加附件标记、
    ; v6.8.14 加页码），写死 y 坐标会与新增行重叠。
    ; v6.8.19 在行距行上方再画一条分隔线，把“字体字号”与“正文行距”分成两个视觉区块。
    editGui.Add("Text", "x415 y" (curY + 4) " w470 h1 BackgroundDDDDDD")
    dlgSpacingLabel := editGui.Add("Text", "x415 y" (curY + 16) " w58", "正文行距:")
    dlgSpacingStar := AddDeviationStar(editGui, "x473 y" (curY + 16) " w10")
    dlgSpacingCtrl := editGui.Add("ComboBox", "x486 y" (curY + 12) " w150", LineSpacingPresets)
    dlgSpacingCtrl.Text := ActiveStyle.Spacing
    dlgSpacingCtrl.OnEvent("Change", UpdateProfileDeviationHighlight)
    editGui.SetFont("s8 c999999")
    editGui.Add("Text", "x646 y" (curY + 16) " w238", "国标：固定 28 磅；含“倍”即非国标")
    editGui.SetFont("s9 c000000")

    ; --- 页眉设置（通栏，v6.8.20 新增）---
    ; 本版只做界面与方案保存，排版时还不会写 Word 页眉（见文件头 v6.8.20 说明）。
    ; 四列栅格与上方两列栅格对齐：列起点 32 / 246 / 460 / 674（标签宽 60），
    ; 控件起点 96 / 310 / 524 / 738（控件宽 134），列间留 16 像素；
    ; 两行 y600 / y634，与「排版选项」「字体字号」两个组框底边（y564）留 12 像素。
    ; 组框 880 宽（20–900），正好等于左列 360 + 间距 20 + 右列 500；
    ; 底边 676，按钮行 690，窗口 h744。
    editGui.Add("GroupBox", "x20 y576 w880 h100", "页眉设置")

    ; 处理方式：三态。下拉第 1/2/3 项对应 HeaderMode 的 0/1/2，
    ; 只有第 2 项「添加·替换」才启用右边这些参数，见 UpdateHeaderControlState()。
    editGui.Add("Text", "x32 y604 w60", "处理方式:")
    headerModeIdx := (Integer(HeaderMode) >= 0 && Integer(HeaderMode) <= 2) ? Integer(HeaderMode) + 1 : 1
    dlgHeaderMode := editGui.Add("DropDownList", "x96 y600 w134 Choose" headerModeIdx,
        ["不改变", "添加·替换", "清空"])
    dlgHeaderMode.OnEvent("Change", OnDialogHeaderModeChange)

    ; 页眉文字：留空是合法值（= 用文档文件名）；
    ; EM_SETCUEBANNER 让系统 Edit 自己画灰底，不额外占控件。
    ; 宽度由 350 收窄到 280，给右侧“记住文字”复选框留位置。
    editGui.Add("Text", "x246 y604 w60", "页眉文字:")
    dlgHeaderText := editGui.Add("Edit", "x310 y600 w250 Limit100", HeaderText)
    try SendMessage(0x1501, 1, StrPtr("留空则使用文件名作为页眉"), , "ahk_id " dlgHeaderText.Hwnd)

    ; v6.8.26：只有勾选才把页眉文字写入注册表，详见文件头 v6.8.26 说明。
    ; 默认不勾选，保证下一次排版/切换方案时不会出现上一次的单位名，
    ; 也不让右键菜单/命令行排版复用旧文字。
    dlgHeaderKeep := editGui.Add("Checkbox", "x570 y605 w80 h18 Checked" HeaderKeep, "记住文字")
    ; v6.8.27："记住文字"用红色显示，提醒用户勾选后页眉文字会随方案一起保存并在批量排版时复用；
    ; 其它控件仍保持黑色，只有这一项变色。
    dlgHeaderKeep.SetFont("cFF0000")

    editGui.Add("Text", "x674 y604 w60", "对齐方式:")
    headerAlignIdx := (HeaderAlign == "居左") ? 1 : ((HeaderAlign == "居右") ? 3 : 2)
    dlgHeaderAlign := editGui.Add("DropDownList", "x738 y600 w134 Choose" headerAlignIdx,
        ["居左", "居中", "居右"])

    ; 第二行：字体 / 字号 / 加粗 / 距上边
    editGui.Add("Text", "x32 y638 w60", "字体:")
    dlgHeaderFont := editGui.Add("ComboBox", "x96 y634 w134", WpsFontNameList)
    dlgHeaderFont.Text := HeaderFont

    editGui.Add("Text", "x246 y638 w60", "字号:")
    dlgHeaderSize := editGui.Add("ComboBox", "x310 y634 w134", WpsFontSizeList)
    dlgHeaderSize.Text := HeaderSize

    ; v6.8.25：“加粗”复选框紧跟“加粗：”三个字（约 41 px），
    ; 从 x524 移到 x502；宽度 18，结束于 520。
    editGui.Add("Text", "x460 y638 w36", "加粗：")
    dlgHeaderBold := editGui.Add("Checkbox", "x502 y638 w18 h18 Checked" HeaderBold)

    ; 页眉横线（v6.8.22；v6.8.23 右移；v6.8.24 改名；v6.8.25 回移）：
    ; 不再单独加“页眉横线:”标签，复选框自带文字即可。
    ; 加粗控件结束于 520，页眉横线从 534 开始（保留 14 像素间距）；
    ; 文字结束于约 605，仍远在“距上边”标签（674）左侧。
    dlgHeaderUnderline := editGui.Add("Checkbox", "x534 y638 w88 h18 Checked" HeaderUnderline, "页眉横线")

    editGui.Add("Text", "x674 y638 w60", "距上边:")
    dlgHeaderMargin := editGui.Add("Edit", "x738 y634 w80 Center Number", FmtMargin(HeaderMarginTop))
    ; MakeSpin() 必须紧跟对应的 Edit 创建（不写 x/y 时自动认上一个控件为 buddy），
    ; 传当前值让箭头有个合理起点，函数内部再把 Edit 文本恢复成精确原值。
    dlgHeaderMarginUD := MakeSpin(editGui, dlgHeaderMargin, "Range0-100", HeaderMarginTop)
    editGui.Add("Text", "x824 y638 w20", "mm")

    ; --- 底部按钮（右端与右侧组框右边界 900 对齐）---
    ; v6.8.19：按钮统一高 34，底边距 20；取消与保存间距 12，保存右端严格对齐 900；
    ; “删除此方案”仍在最左，与组框左边界对齐，中间留白不再被按钮挤成一条缝。
    ; v6.8.20：页眉组占用了 y576–676，按钮行由 y578 下移到 y690，窗口 h632 → h744。
    btnDelete := editGui.Add("Button", "x20 y690 w120 h34", "删除此方案")
    btnDelete.OnEvent("Click", DeleteActiveProfileFromEditor)
    btnCancel := editGui.Add("Button", "x656 y690 w112 h34", "取消")
    btnCancel.OnEvent("Click", (*) => CloseProfileEditor())
    btnSave := editGui.Add("Button", "x780 y690 w120 h34 Default", "保存")
    btnSave.OnEvent("Click", SaveProfileEditor)

    ; 边距输入与微调
    dlgTop.OnEvent("Change", OnDialogMarginEditChange)
    dlgBottom.OnEvent("Change", OnDialogMarginEditChange)
    dlgLeft.OnEvent("Change", OnDialogMarginEditChange)
    dlgRight.OnEvent("Change", OnDialogMarginEditChange)
    dlgTopUD.OnEvent("Change", (ctrl, *) => AdjustDialogMarginBySpinner(dlgTop, ctrl))
    dlgBottomUD.OnEvent("Change", (ctrl, *) => AdjustDialogMarginBySpinner(dlgBottom, ctrl))
    dlgLeftUD.OnEvent("Change", (ctrl, *) => AdjustDialogMarginBySpinner(dlgLeft, ctrl))
    dlgRightUD.OnEvent("Change", (ctrl, *) => AdjustDialogMarginBySpinner(dlgRight, ctrl))
    dlgTop.OnEvent("Focus", EditSelectAll)
    dlgBottom.OnEvent("Focus", EditSelectAll)
    dlgLeft.OnEvent("Focus", EditSelectAll)
    dlgRight.OnEvent("Focus", EditSelectAll)

    ; 打开窗口时按当前参数初始化星号显示
    UpdateProfileDeviationHighlight()
    ; 按当前网格类型启用/禁用“每页行数”“每行字符数”两个输入框
    UpdateGridControlState()
    ; v6.8.20：按页眉“处理方式”启用/禁用页眉参数（默认「不改变」→ 全部置灰）
    UpdateHeaderControlState()

    ; 编辑期间锁定主窗口，避免同时切换方案
    mainGui.Opt("+Disabled")
    editGui.Show("w920 h744")

    ; 开始检测鼠标是否停在三个排版选项上，用于显示详细说明
    SetTimer(UpdateOptionTooltip, 150)
}

; =====================================================================
; [文档网格：编辑窗口联动]
; =====================================================================
; 下拉框顺序固定为：1=指定行和字符网格、2=只指定行网格、3=不指定网格。
; 下标与 Word 的 LayoutMode 枚举（1/2/0）不是一一对应，所以统一用下面这个映射，
; 改下拉文案时不必改代码，但**不要调整顺序**。
OnDialogGridChange(*) {
    global isProgramUpdating
    if (isProgramUpdating)
        return
    ; 用户第一次切到某种网格类型时，把“该类型的推荐值”补进空着的输入框，
    ; 否则会出现“只指定行网格但行数是 0 / 空”这种没法排版的组合。
    ApplyGridDefaultsForMode()
    UpdateGridControlState()
}

; 切换网格类型时补齐推荐值（只在“看起来没设过”时才动，不覆盖用户手工改的值）：
;   指定行和字符网格（1）：行数无效 → 22，字符数无效 → 28（国标口径）
;   只指定行网格（2）：行数无效，或仍等于国标的 22（说明用户没为这种类型单独设过）
;                      → 44（Word 默认口径）
;   不指定网格（0）：什么都不动，值留着，切回来还能直接用
; 判定“无效”用 IsNumber + 大于 0，因为输入框允许被清空。
ApplyGridDefaultsForMode() {
    global dlgGridLines, dlgGridChars
    if (!IsSet(dlgGridLines) || !IsObject(dlgGridLines))
        return

    mode := GetGridModeFromDropdown()
    linesEmpty := !IsNumber(dlgGridLines.Value) || Integer(dlgGridLines.Value) < 1
    charsEmpty := !IsNumber(dlgGridChars.Value) || Integer(dlgGridChars.Value) < 1

    if (mode == 1) {
        if (linesEmpty)
            dlgGridLines.Value := 22
        if (charsEmpty)
            dlgGridChars.Value := 28
    } else if (mode == 2) {
        if (linesEmpty || Integer(dlgGridLines.Value) == 22)
            dlgGridLines.Value := 44
    }
}

; 按当前网格类型启用/禁用行数、字符数输入框：
;   不指定网格      → 两个都禁用
;   只指定行网格    → 只禁用“每行字符数”
;   指定行和字符网格 → 两个都可用
UpdateGridControlState() {
    global dlgGridMode, dlgGridLines, dlgGridChars, dlgGridLinesUD, dlgGridCharsUD
    if (!IsSet(dlgGridMode) || !IsObject(dlgGridMode))
        return

    mode := GetGridModeFromDropdown()
    linesOn := (mode == 1 || mode == 2)
    charsOn := (mode == 1)

    try dlgGridLines.Enabled := linesOn
    try dlgGridLinesUD.Enabled := linesOn
    try dlgGridChars.Enabled := charsOn
    try dlgGridCharsUD.Enabled := charsOn
}

; 下拉框下标 → GridMode 值（1/2/0）
GetGridModeFromDropdown() {
    global dlgGridMode
    switch dlgGridMode.Value {
        case 1: return 1
        case 2: return 2
        default: return 0      ; 3 = 不指定网格；空值也按不指定处理
    }
}

; 从编辑窗口读出网格的三个值，返回 [GridMode, GridLines, GridChars]
GetGridFromEditor() {
    global dlgGridLines, dlgGridChars
    return [GetGridModeFromDropdown(),
            IsNumber(dlgGridLines.Value) ? Integer(dlgGridLines.Value) : 0,
            IsNumber(dlgGridChars.Value) ? Integer(dlgGridChars.Value) : 0]
}

; 把 [GridMode, GridLines, GridChars] 写回编辑窗口控件
SetGridToEditor(arr) {
    global dlgGridMode, dlgGridLines, dlgGridChars
    dlgGridMode.Value := (arr[1] == 1) ? 1 : ((arr[1] == 2) ? 2 : 3)
    dlgGridLines.Value := arr[2]
    dlgGridChars.Value := arr[3]
    UpdateGridControlState()
}

; =====================================================================
; [页眉：编辑窗口联动（v6.8.20）]
; =====================================================================
; 页眉“处理方式”三态与控件的启用关系：
;   1 不改变    → 保留文档原有页眉，六个页眉参数一个都不生效 → 全部置灰
;   2 添加·替换 → 全部可用
;   3 清空      → 只删除原有页眉，参数同样不生效 → 全部置灰
; 置灰列表里必须包含 dlgHeaderMarginUD（数字框的上下箭头），
; 否则会出现“输入框灰了、箭头还能点”的怪状态。
; 与 UpdateGridControlState() 一样，先判断控件是否已创建：
; 窗口关闭后全局变量会被置空，残留的定时器/事件调用不能报错。
UpdateHeaderControlState() {
    global dlgHeaderMode, dlgHeaderText, dlgHeaderAlign
    global dlgHeaderFont, dlgHeaderSize, dlgHeaderBold, dlgHeaderUnderline
    global dlgHeaderMargin, dlgHeaderMarginUD
    if (!IsSet(dlgHeaderMode) || !IsObject(dlgHeaderMode))
        return

    on := (dlgHeaderMode.Value == 2)      ; 下拉第二项 = 添加·替换
    for ctrl in [dlgHeaderText, dlgHeaderAlign, dlgHeaderFont, dlgHeaderSize,
                 dlgHeaderBold, dlgHeaderUnderline, dlgHeaderMargin, dlgHeaderMarginUD] {
        if (IsObject(ctrl))
            try ctrl.Enabled := on
    }
}

; 用户切换“处理方式”时重新计算置灰状态。
; isProgramUpdating 期间（程序自己改下拉值）不响应，避免和初始化流程打架。
OnDialogHeaderModeChange(*) {
    global isProgramUpdating
    if (isProgramUpdating)
        return
    UpdateHeaderControlState()
}

; 返回某个边距预设对应的推荐网格 [GridMode, GridLines, GridChars]；
; “用户自定义”没有推荐值，返回空字符串（调用方据此跳过联动）。
GetGridPresetRecommendation(preset) {
    if (preset == "党政机关公文格式")
        return [1, 22, 28]          ; GB/T 9704-2012：每页 22 行、每行 28 个汉字
    if (preset == "WORD默认边距" || preset == "WORD默认页边距")
        return [2, 44, 0]           ; Word 新建文档：只指定行网格、44 行、每行字数不指定
    return ""
}

; 切换“边距预设”时联动网格。
; 只在网格仍等于“另一套预设的推荐值”时才跟随，说明用户没有手工改过网格；
; 一旦手工改过（例如故意设成 30 行），切边距预设不会覆盖用户的设置。
SyncGridWithMarginPreset(preset) {
    global isProgramUpdating, dlgGridMode
    if (!IsSet(dlgGridMode) || !IsObject(dlgGridMode))
        return

    target := GetGridPresetRecommendation(preset)
    if (!IsObject(target))
        return                  ; 用户自定义：不联动
    other := GetGridPresetRecommendation(preset == "党政机关公文格式"
        ? "WORD默认边距" : "党政机关公文格式")
    if (!IsObject(other))
        return

    cur := GetGridFromEditor()
    if (cur[1] == other[1] && cur[2] == other[2] && cur[3] == other[3]) {
        isProgramUpdating := true
        SetGridToEditor(target)
        isProgramUpdating := false
    }
    UpdateGridControlState()
}

; 返回排版选项的说明文字（1=表格保护，2=防孤行，3=附件另起一页）
GetOptionTooltipText(optId) {
    switch optId {
        case 1:
            return "勾选：保留表格原有的字体、字号、缩进、`n对齐和行距，排版时不改动表格。`n不勾选：按方案里的“表格内容”统一表格`n字体字号，并整体居中、文字左对齐。"
        case 2:
            return "标题落在页面底部、正文被挤到下一页时，`n自动把标题和下一段一起移到下一页，`n避免标题单独留在页底（孤行）。"
        case 3:
            return "勾选：给每个“附件1”“附件2”标记设置段前分页，`n让附件另面编排（GB/T 9704-2012 7.3.8）。`n标记本身已带分页符时不会再重复加页。"
    }
    return ""
}

; 判断鼠标是否停在三个排版选项上，返回 1/2/3，不在选项上时返回 0
; 说明：WindowFromPoint 会命中覆盖在其上方的 GroupBox，所以这里按控件矩形判断
GetOptionUnderMouse(mx, my) {
    global editGui, dlgProtectTable, dlgPreventOrphan, dlgAttachNewPage
    if (!IsSet(dlgProtectTable) || !IsObject(dlgProtectTable))
        return 0
    try
        WinGetClientPos(&ox, &oy, , , "ahk_id " editGui.Hwnd)
    catch
        return 0

    opts := [dlgProtectTable, dlgPreventOrphan, dlgAttachNewPage]
    for idx, ctl in opts {
        try
            ctl.GetPos(&cx, &cy, &cw, &ch)
        catch
            continue
        ; 上下各放宽 4 像素，鼠标停在文字附近也能触发
        if (mx >= ox + cx && mx <= ox + cx + cw && my >= oy + cy - 4 && my <= oy + cy + ch + 4)
            return idx
    }
    return 0
}

; 隐藏悬停说明
HideOptionTooltip() {
    global optTipId
    if (optTipId != 0) {
        ToolTip(, , , 1)
        optTipId := 0
    }
}

; 定时检测光标位置并显示/隐藏三个选项的悬停说明
UpdateOptionTooltip(*) {
    global editGui, optTipId
    if (!IsSet(editGui) || !IsObject(editGui) || !WinActive("ahk_id " editGui.Hwnd)) {
        HideOptionTooltip()
        return
    }

    MouseGetPos(&mx, &my)
    optId := GetOptionUnderMouse(mx, my)
    if (optId == 0) {
        HideOptionTooltip()
        return
    }
    if (optTipId == optId)
        return

    optTipId := optId
    tipX := mx + 20
    if (tipX > A_ScreenWidth - 420)
        tipX := A_ScreenWidth - 420
    tipY := my + 24
    if (tipY > A_ScreenHeight - 140)
        tipY := my - 100
    ToolTip(GetOptionTooltipText(optId), tipX, tipY, 1)
}

CloseProfileEditor(*) {
    global mainGui, editGui, editorOriginalName
    global dlgName, dlgMarginType, dlgTop, dlgBottom, dlgLeft, dlgRight
    global dlgLblMarginType, dlgLblTop, dlgLblBottom, dlgLblLeft, dlgLblRight
    global dlgStarMarginType, dlgStarTop, dlgStarBottom, dlgStarLeft, dlgStarRight
    global dlgTopUD, dlgBottomUD, dlgLeftUD, dlgRightUD
    global dlgPageFormat, dlgSkipSingle
    global dlgProtectTable, dlgPreventOrphan, dlgAttachNewPage, dlgMaxHeadingLen
    global dlgFontRows, dlgSpacingCtrl, dlgSpacingStar, dlgSpacingLabel
    ; v6.8.20 页眉组的 10 个控件（含数字框的 UpDown；v6.8.26 增加记住文字），关闭窗口时必须一起置空，
    ; 否则下次打开编辑器会残留上一轮的控件对象，UpdateHeaderControlState() 可能改到已销毁的控件。
    global dlgHeaderMode, dlgHeaderText, dlgHeaderAlign
    global dlgHeaderFont, dlgHeaderSize, dlgHeaderBold, dlgHeaderUnderline
    global dlgHeaderMargin, dlgHeaderMarginUD, dlgHeaderKeep
    global optTipId

    ; 停止悬停说明检测并隐藏提示
    SetTimer(UpdateOptionTooltip, 0)
    ToolTip(, , , 1)
    optTipId := 0

    if (IsSet(editGui) && IsObject(editGui)) {
        try editGui.Destroy()
    }
    editGui := ""
    dlgName := "", dlgMarginType := "", dlgTop := "", dlgBottom := "", dlgLeft := "", dlgRight := ""
    dlgLblMarginType := "", dlgLblTop := "", dlgLblBottom := ""
    dlgLblLeft := "", dlgLblRight := ""
    dlgStarMarginType := "", dlgStarTop := "", dlgStarBottom := ""
    dlgStarLeft := "", dlgStarRight := ""
    dlgTopUD := "", dlgBottomUD := "", dlgLeftUD := "", dlgRightUD := ""
    dlgPageFormat := "", dlgSkipSingle := ""
    dlgProtectTable := "", dlgPreventOrphan := "", dlgAttachNewPage := "", dlgMaxHeadingLen := ""
    dlgFontRows := [], dlgSpacingCtrl := "", dlgSpacingStar := "", dlgSpacingLabel := ""
    dlgGridMode := "", dlgGridLines := "", dlgGridChars := ""
    dlgGridLinesUD := "", dlgGridCharsUD := ""
    dlgHeaderMode := "", dlgHeaderText := "", dlgHeaderAlign := ""
    dlgHeaderFont := "", dlgHeaderSize := "", dlgHeaderBold := "", dlgHeaderUnderline := ""
    dlgHeaderMargin := "", dlgHeaderMarginUD := "", dlgHeaderKeep := ""
    editorOriginalName := ""

    if (IsSet(mainGui) && IsObject(mainGui))
        mainGui.Opt("-Disabled")
}

OnDialogMarginTypeChange(*) {
    global isProgramUpdating, dlgMarginType, dlgTop, dlgBottom, dlgLeft, dlgRight
    global dlgCustomTop, dlgCustomBottom, dlgCustomLeft, dlgCustomRight
    if (isProgramUpdating)
        return

    isProgramUpdating := true
    sel := dlgMarginType.Text
    if (sel == "党政机关公文格式") {
        dlgTop.Value := "37", dlgBottom.Value := "35", dlgLeft.Value := "28", dlgRight.Value := "26"
    } else if (sel == "WORD默认边距" || sel == "WORD默认页边距") {
        dlgTop.Value := "25.4", dlgBottom.Value := "25.4", dlgLeft.Value := "31.8", dlgRight.Value := "31.8"
    } else if (sel == "用户自定义") {
        dlgTop.Value := FmtMargin(dlgCustomTop)
        dlgBottom.Value := FmtMargin(dlgCustomBottom)
        dlgLeft.Value := FmtMargin(dlgCustomLeft)
        dlgRight.Value := FmtMargin(dlgCustomRight)
    }
    isProgramUpdating := false
    ; 边距预设变了，网格跟着推荐值联动（手工改过网格则保持不动）
    SyncGridWithMarginPreset(sel)
    UpdateProfileDeviationHighlight()
}

OnDialogMarginEditChange(*) {
    global isProgramUpdating, dlgMarginType, dlgTop, dlgBottom, dlgLeft, dlgRight
    global dlgCustomTop, dlgCustomBottom, dlgCustomLeft, dlgCustomRight
    if (isProgramUpdating)
        return

    isProgramUpdating := true
    try dlgMarginType.Choose("用户自定义")
    isProgramUpdating := false

    dlgCustomTop    := IsNumber(dlgTop.Value) ? Float(dlgTop.Value) : 0.0
    dlgCustomBottom := IsNumber(dlgBottom.Value) ? Float(dlgBottom.Value) : 0.0
    dlgCustomLeft   := IsNumber(dlgLeft.Value) ? Float(dlgLeft.Value) : 0.0
    dlgCustomRight  := IsNumber(dlgRight.Value) ? Float(dlgRight.Value) : 0.0

    UpdateProfileDeviationHighlight()
}

AdjustDialogMarginBySpinner(editCtrl, udCtrl) {
    global isProgramUpdating
    if (isProgramUpdating)
        return

    diff := udCtrl.Value - 50
    isProgramUpdating := true
    udCtrl.Value := 50
    isProgramUpdating := false

    currentVal := IsNumber(editCtrl.Value) ? Float(editCtrl.Value) : 0.0
    newVal := Max(0.0, Round(currentVal + diff * 1, 1))

    isProgramUpdating := true
    editCtrl.Value := FmtMargin(newVal)
    isProgramUpdating := false

    OnDialogMarginEditChange()
}

SaveProfileEditor(*) {
    global ActiveProfile, editorOriginalName
    global dlgName, dlgMarginType, dlgTop, dlgBottom, dlgLeft, dlgRight
    global dlgCustomTop, dlgCustomBottom, dlgCustomLeft, dlgCustomRight
    global dlgPageFormat, dlgSkipSingle
    global dlgProtectTable, dlgPreventOrphan, dlgAttachNewPage, dlgMaxHeadingLen
    global dlgFontRows, dlgSpacingCtrl
    ; v6.8.20 页眉组控件（含数字框的 UpDown；v6.8.22 增加页眉横线复选框；v6.8.26 增加记住文字）
    global dlgHeaderMode, dlgHeaderText, dlgHeaderAlign
    global dlgHeaderFont, dlgHeaderSize, dlgHeaderBold, dlgHeaderUnderline
    global dlgHeaderMargin, dlgHeaderMarginUD, dlgHeaderKeep

    newName := Trim(dlgName.Value)
    if (!ValidateProfileName(newName))
        return
    if (IsBuiltinProfile(newName)) {
        MsgBox("“" GetBuiltinProfileName() "”为内置只读方案，请换一个名称。", "提示", "Icon!")
        return
    }

    ; 校验边距必须为数字
    for item in [["上", dlgTop], ["下", dlgBottom], ["左", dlgLeft], ["右", dlgRight]] {
        if (!IsNumber(item[2].Value)) {
            MsgBox("边距“" item[1] "”必须是数字。", "提示", "Icon!")
            return
        }
    }

    ; 校验标题阈值
    lenVal := dlgMaxHeadingLen.Value
    if (!IsNumber(lenVal) || Integer(lenVal) < 10 || Integer(lenVal) > 200) {
        MsgBox("标题阈值必须是 10 到 200 之间的整数。", "提示", "Icon!")
        return
    }

    ; 校验文档网格：行数最少 1（写 0 会让 Word 的 COM 直接崩溃），
    ; 每行字符数 0 表示“不指定”，只在“指定行和字符网格”下生效。
    gridArr := GetGridFromEditor()
    if (gridArr[2] < 1 || gridArr[2] > 100) {
        MsgBox("每页行数必须是 1 到 100 之间的整数。`n`n提示：Word 的“每页行数”不能为 0。", "提示", "Icon!")
        return
    }
    if (gridArr[3] < 0 || gridArr[3] > 100) {
        MsgBox("每行字符数必须是 0 到 100 之间的整数。`n`n0 表示不指定每行字符数。", "提示", "Icon!")
        return
    }

    ; 校验页眉参数（v6.8.20）。下拉第 1/2/3 项对应 HeaderMode 的 0/1/2，
    ; 存语义值而不是下拉下标，改文案/顺序不影响已保存的方案。
    headerModeVal := Integer(dlgHeaderMode.Value) - 1
    if (headerModeVal < 0 || headerModeVal > 2)
        headerModeVal := 0
    ; 页眉文字：留空是合法值（= 用文档文件名），所以这里**不能**用空串判错，
    ; 只去掉首尾空白，避免用户误敲空格变成“有一个空格的页眉”。
    headerTextVal := Trim(dlgHeaderText.Value)
    ; 对齐/字体/字号理论上不会为空（都是下拉），但 ComboBox 允许手工输入，
    ; 空值时回落到内置默认，免得方案里存进空字符串让后续 COM 拿到非法字体名。
    headerAlignVal := Trim(dlgHeaderAlign.Text)
    if (headerAlignVal == "")
        headerAlignVal := "居中"
    headerFontVal := Trim(dlgHeaderFont.Text)
    if (headerFontVal == "")
        headerFontVal := "宋体"
    headerSizeVal := Trim(dlgHeaderSize.Text)
    if (headerSizeVal == "")
        headerSizeVal := "四号"
    headerBoldVal := dlgHeaderBold.Value ? 1 : 0
    headerUnderlineVal := dlgHeaderUnderline.Value ? 1 : 0
    ; v6.8.26：HeaderKeep=0 时，发送空字符串进入方案，不保存隐私文字；
    ; 勾选“记住文字”才保留原始 HeaderText。
    headerKeepVal := dlgHeaderKeep.Value ? 1 : 0
    headerTextVal := headerKeepVal ? headerTextVal : ""
    ; 距上边：只在“添加·替换”时校验。另外两种模式下这个输入框是灰的，
    ; 若也强制校验，用户切到“清空”后就没法保存方案了。
    headerMarginRaw := dlgHeaderMargin.Value
    if (headerModeVal == 1) {
        if (!IsNumber(headerMarginRaw) || Float(headerMarginRaw) < 0 || Float(headerMarginRaw) > 100) {
            MsgBox("页眉“距上边”必须是 0 到 100 之间的数字（单位：毫米）。", "提示", "Icon!")
            return
        }
    }
    ; 非数字时回落到默认 15 mm，不把 0 写进方案（0 会把页眉压到页面最顶端）
    headerMarginVal := IsNumber(headerMarginRaw) ? FmtMargin(headerMarginRaw) : 15

    oldName := editorOriginalName
    if (newName != oldName && ProfileExists(newName)) {
        if (MsgBox("方案“" newName "”已存在，是否覆盖？", "确认覆盖", "YesNo Icon?") != "Yes")
            return
    }

    obj := {
        MarginType: dlgMarginType.Value,
        MarginTop: dlgTop.Value, MarginBottom: dlgBottom.Value,
        MarginLeft: dlgLeft.Value, MarginRight: dlgRight.Value,
        CustomTop: dlgCustomTop, CustomBottom: dlgCustomBottom,
        CustomLeft: dlgCustomLeft, CustomRight: dlgCustomRight,
        PageNumFormat: dlgPageFormat.Text,
        SkipSinglePageNum: dlgSkipSingle.Value,
        GridMode: gridArr[1], GridLines: gridArr[2], GridChars: gridArr[3],
        CurrentFontScheme: GetBuiltinProfileName(),
        OptProtectTable: dlgProtectTable.Value,
        OptPreventOrphan: dlgPreventOrphan.Value,
        OptAttachmentNewPage: dlgAttachNewPage.Value,
        MaxHeadingLen: Integer(dlgMaxHeadingLen.Value),
        ; v6.8.20 页眉参数（v6.8.26 起仅勾选“记住文字”才保存 HeaderText）；v6.8.22 增加页眉横线
        HeaderMode: headerModeVal, HeaderText: headerTextVal, HeaderAlign: headerAlignVal,
        HeaderMarginTop: headerMarginVal, HeaderFont: headerFontVal,
        HeaderSize: headerSizeVal, HeaderBold: headerBoldVal,
        HeaderUnderline: headerUnderlineVal, HeaderKeep: headerKeepVal
    }
    ; 字体字号与边距等一起保存进方案
    for row in dlgFontRows {
        fontProp := row.prefix . "Font"
        sizeProp := row.prefix . "Size"
        boldProp := row.prefix . "Bold"
        obj.%fontProp% := Trim(row.font.Text)
        obj.%sizeProp% := Trim(row.size.Text)
        obj.%boldProp% := Integer(row.bold.Value)
    }
    obj.Spacing := Trim(dlgSpacingCtrl.Text)

    ProfileWriteObject(newName, obj)

    ; 重命名时删除旧方案键；覆盖同名方案时无需删除
    if (oldName != "" && oldName != newName && !IsBuiltinProfile(oldName)) {
        try RegDeleteKey(REG_PROFILES_PATH "\" oldName)
    }

    ActiveProfile := newName
    ApplyProfile(newName, false)
    ShowFloatingToast("方案已保存", "当前方案：" newName, 2000, true)
    CloseProfileEditor()
}

DeleteActiveProfileFromEditor(*) {
    global ActiveProfile, editorOriginalName
    name := (editorOriginalName != "") ? editorOriginalName : ActiveProfile
    if (IsBuiltinProfile(name))
        return

    if (MsgBox("确定要删除方案“" name "”吗？`n`n此操作不可撤销。", "删除方案", "YesNo Icon?") != "Yes")
        return

    try RegDeleteKey(REG_PROFILES_PATH "\" name)

    names := GetUserProfileNames()
    newActive := names.Length > 0 ? names[1] : GetBuiltinProfileName()
    ActiveProfile := newActive
    ApplyProfile(newActive)
    ShowFloatingToast("方案已删除", "已切换到：" newActive, 2200, true)
    CloseProfileEditor()
}

RestoreSelectedLVItem(*) {
    global uiLV, uiSB, GuiBatchRunning
    ; 排版途中不要动列表/文件：行号正在被 Loop count 使用
    if (GuiBatchRunning) {
        uiSB.SetText(" 正在排版，请等本批处理结束后再恢复备份。")
        return
    }
    row := uiLV.GetNext(0)
    if (row > 0) {
        path := uiLV.GetText(row, 4)
        if (RestoreLatestBackup(path))
            uiLV.Modify(row, "Col2", "已恢复备份")
    } else {
        uiSB.SetText(" 请先在列表里选中一个文件。")
    }
}

; 双击任务列表中的文件，用系统默认程序打开
OpenSelectedLVFile(*) {
    global uiLV, uiSB
    row := uiLV.GetNext(0)
    if (row <= 0) {
        uiSB.SetText(" 请先在列表里选中一个文件。")
        return
    }
    filePath := uiLV.GetText(row, 4)
    if (filePath == "" || !FileExist(filePath)) {
        uiSB.SetText(" 文件不存在：" filePath)
        ShowFloatingToast("打开失败", "文件不存在或已被移动：`n" filePath, 3000, false)
        return
    }
    try {
        Run('"' filePath '"')
        uiSB.SetText(" 已打开：" filePath)
    } catch Error as err {
        uiSB.SetText(" 打开失败：" err.Message)
        ShowFloatingToast("打开失败", err.Message, 3000, false)
    }
}

; 右键“打开所在目录”（v6.8.22）：打开资源管理器并选中该文件；
; 文件刚好被移走时退回“只打开目录”，不弹模态框。
; 注意：explorer.exe 的参数里带空格时必须整体加引号，否则会被拆成多个参数。
OpenSelectedLVFolder(*) {
    global uiLV, uiSB
    row := uiLV.GetNext(0)
    if (row <= 0) {
        uiSB.SetText(" 请先在列表里选中一个文件。")
        return
    }
    filePath := uiLV.GetText(row, 4)
    SplitPath(filePath, , &dir)
    if (dir == "" || !DirExist(dir)) {
        uiSB.SetText(" 找不到文件所在目录：" filePath)
        return
    }
    try {
        Run('explorer.exe /select,"' filePath '"')
        uiSB.SetText(" 已打开所在目录：" dir)
    } catch {
        try {
            Run('"' dir '"')
            uiSB.SetText(" 已打开所在目录：" dir)
        } catch Error as err {
            uiSB.SetText(" 打开目录失败：" err.Message)
        }
    }
}

; 右键“移除所选文件”（v6.8.22）：只从列表里移除，不动磁盘上的文件。
; 必须从后往前删：从前往后删会让后面的行号整体上移，导致漏删或删错行。
RemoveSelectedLVItems(*) {
    global uiLV, uiSB, GuiBatchRunning
    if (GuiBatchRunning) {
        uiSB.SetText(" 正在排版，请等本批处理结束后再移除列表项。")
        return
    }

    rows := []
    row := 0
    while (row := uiLV.GetNext(row))
        rows.Push(row)

    if (rows.Length == 0) {
        uiSB.SetText(" 请先在列表里选中要移除的文件。")
        return
    }
    while (rows.Length)
        uiLV.Delete(rows.Pop())

    uiSB.SetText(" 已移除所选任务，当前列表共 " uiLV.GetCount() " 个。")
}

CheckButtonHover(wParam, lParam, msg, hwnd) {
    global uiBtnKill, uiBakOpen
    if (IsSet(uiBakOpen) && IsObject(uiBakOpen) && hwnd == uiBakOpen.Hwnd) {
        ToolTip("打开备份文件夹")
        SetTimer(HideButtonToolTip, -1200)
        return
    }
    if (IsSet(uiBtnKill) && IsObject(uiBtnKill) && hwnd == uiBtnKill.Hwnd) {
        ToolTip("强制结束由本程序在后台调用的排版进程`n（安全操作，不影响您自行打开的文档）")
        SetTimer(() => ToolTip(), -300)
    }
}

HideButtonToolTip() {
    ToolTip()
}

GuiClose(*) {
    global BakOpenHIcon
    if (BakOpenHIcon) {
        DllCall("user32\DestroyIcon", "Ptr", BakOpenHIcon)
        BakOpenHIcon := 0
    }
    ExitApp()
}

; 排版前确认备份开关：返回 true 表示可以继续排版
; 未开启备份时弹出“是/否”提醒，默认“是”（开启备份）；关闭或显示失败则视为取消
EnsureBackupConfirmed() {
    global CreateBAK, uiBAK
    if (CreateBAK)
        return true

    bakTip := MsgBox("检测到“排版前生成原始备份”当前处于关闭状态。`n`n"
        . "是否在开始排版前为每个文件生成原始备份？`n`n"
        . "· 选择“是”：开启备份并继续排版（推荐）`n"
        . "· 选择“否”：本次不备份，直接排版", "备份提醒", "YesNo Icon!")

    if (bakTip == "Yes") {
        CreateBAK := 1
        if (IsSet(uiBAK) && IsObject(uiBAK))
            uiBAK.Value := 1
        SetConfigVal("Options", "CreateBAK", CreateBAK)
        return true
    }
    return (bakTip == "No")
}

ClearList(*) {
    global uiLV, uiSB, GuiBatchRunning
    ; v6.8.22：按钮在批量期间已被置灰，但右键菜单绕过了按钮状态，这里必须自己拦
    if (GuiBatchRunning) {
        uiSB.SetText(" 正在排版，请等本批处理结束后再清空列表。")
        return
    }
    uiLV.Delete()
    uiSB.SetText(" 列表已清空。")
}

KillAllWordProcesses(*) {
    global uiSB, ManagedWordPIDs
    killed := 0
    for pid in ManagedWordPIDs {
        if (ProcessExist(pid)) {
            try {
                ProcessClose(pid)
                killed++
                Sleep(100)
            } catch {
                continue
            }
        }
    }
    ManagedWordPIDs := Map()
    if (killed > 0)
        uiSB.SetText(" 清理完毕！共结束了 " killed " 个由本程序在后台调用的驻留进程。")
    else
        uiSB.SetText(" 当前系统未检测到由本程序调用的后台排版进程。")
}

EditSelectAll(ctrlObj, *) {
    SendMessage(0x00B1, 0, -1, ctrlObj.Hwnd)
}

GuiDropFiles(guiObj, ctrlObj, fileArray, x, y) {
    global uiLV, uiSB
    existingPaths := Map()
    Loop uiLV.GetCount() {
        p := uiLV.GetText(A_Index, 4)
        if (p != "")
            existingPaths[Format("{:L}", p)] := true
    }

    ; 拖入的文件夹若包含子文件夹，先询问是否连同子文件夹一起扫描；默认按钮为“是”
    scanSubfolders := false
    for path in fileArray {
        if (InStr(FileExist(path), "D") && HasSubFolders(path)) {
            answer := MsgBox("检测到拖入的文件夹中包含子文件夹。`n`n"
                . "是否扫描子文件夹中的 Word 文件？`n`n"
                . "· 选择“是”：包含所有子文件夹中的文件（推荐）`n"
                . "· 选择“否”：只添加该文件夹第一层的文件", "扫描子文件夹", "YesNo Icon?")
            scanSubfolders := (answer == "Yes")
            break
        }
    }

    addedCount := 0, skippedCount := 0
    for path in fileArray {
        if InStr(FileExist(path), "D") {
            ; FR = 文件 + 递归子文件夹；F = 仅当前文件夹第一层
            fileOpts := scanSubfolders ? "FR" : "F"
            Loop Files, path "\*.*", fileOpts {
                ext := Format("{:L}", A_LoopFileExt)
                if ((ext == "doc" || ext == "docx") && SubStr(A_LoopFileName, 1, 2) != "~$") {
                    fullPathKey := Format("{:L}", A_LoopFilePath)
                    if (!existingPaths.Has(fullPathKey)) {
                        existingPaths[fullPathKey] := true
                        ; v6.8.25：第三列显示文件夹；第四列隐藏完整路径供按行读取
                        SplitPath(A_LoopFilePath, , &parentDir)
                        uiLV.Add("", A_LoopFileName, "等待排版", parentDir, A_LoopFilePath)
                        addedCount++
                    } else {
                        skippedCount++
                    }
                }
            }
        } else {
            SplitPath(path, &fileName, &parentDir, &fileExt)
            ext := Format("{:L}", fileExt)
            if ((ext == "doc" || ext == "docx") && SubStr(fileName, 1, 2) != "~$") {
                fullPathKey := Format("{:L}", path)
                if (!existingPaths.Has(fullPathKey)) {
                    existingPaths[fullPathKey] := true
                    ; v6.8.25：第三列显示文件夹；第四列隐藏完整路径供按行读取
                    uiLV.Add("", fileName, "等待排版", parentDir, path)
                    addedCount++
                } else {
                    skippedCount++
                }
            }
        }
    }

    statusMsg := " 已添加 " addedCount " 个新任务。"
    if (skippedCount > 0)
        statusMsg .= " (已跳过 " skippedCount " 个重复文件)"
    statusMsg .= " 当前列表共 " uiLV.GetCount() " 个排版任务。"
    uiSB.SetText(statusMsg)
}

; 判断指定目录下是否存在子文件夹
HasSubFolders(dirPath) {
    Loop Files, dirPath "\*", "D"
        return true
    return false
}

SelectBakPath(*) {
    global uiBakPath
    selectedDir := BrowseFolderWithNewButton(uiBakPath.Value)
    if (selectedDir != "") {
        uiBakPath.Value := selectedDir
        SaveSingleSetting("Options", "BakDir", selectedDir)
    }
}

BrowseFolderWithNewButton(initialDir := "") {
    global DefaultBakDir
    ; 起始路径为空或已不存在时回退到“文档”，避免对话框定位失败
    startDir := Trim(initialDir)
    if (startDir == "")
        startDir := DefaultBakDir
    if (!DirExist(startDir))
        startDir := DirExist(A_MyDocuments) ? A_MyDocuments : ""

    ; 前缀 "*" 允许向上逐级导航（旧实现把路径当作根目录，导致无法切换到其它目录）；
    ; Options=3 同时提供“新建文件夹”按钮和可输入的编辑框
    try return DirSelect(startDir = "" ? "" : "*" startDir, 3, "请选择备份存放文件夹")
    catch
        return ""
}

; 在资源管理器中打开当前设置的备份目录
OpenBakFolder(*) {
    global uiBakPath, DefaultBakDir
    dir := Trim(uiBakPath.Value)
    if (dir == "")
        dir := DefaultBakDir
    if (!DirExist(dir)) {
        ; 目录尚未创建时自动创建，便于用户提前查看或管理备份位置
        try DirCreate(dir)
        catch {
            ShowFloatingToast("打开失败", "备份目录不存在且无法创建：`n" dir, 3000, false)
            return
        }
    }
    try Run('"' dir '"')
    catch {
        ShowFloatingToast("打开失败", "无法打开备份目录，请检查路径或权限！", 3000, false)
    }
}

; =====================================================================
; [系统标准图标：取 shell 的 SIID_FOLDEROPEN，并贴到方形图标按钮上]
; =====================================================================
GetStockIconHandle(iconID, flags) {
    ; SHSTOCKICONINFO：cbSize + 对齐 + hIcon + iSysImageIndex + iIcon + szPath[260]
    static infoSize := 4 + A_PtrSize + 4 + 4 + 520 + (A_PtrSize = 8 ? 4 : 0)
    info := Buffer(infoSize, 0)
    NumPut("UInt", infoSize, info, 0)
    if (DllCall("shell32\SHGetStockIconInfo", "Int", iconID, "UInt", flags, "Ptr", info, "Int") != 0)
        return 0
    return NumGet(info, A_PtrSize = 8 ? 8 : 4, "Ptr")
}

SetButtonStockIcon(btn, iconID := 4) {
    global BakOpenHIcon
    ; SHGSI_ICON(0x100) | SHGSI_SMALLICON(0x1)
    hIcon := GetStockIconHandle(iconID, 0x101)
    if (!hIcon)
        return false

    ; BS_ICON 按钮：BM_SETIMAGE(0x00F7) + IMAGE_ICON(1) 设置图标；
    ; BM_SETIMAGE 返回“之前的图标句柄”，首次设置返回 0 属正常，故不据此判断成败。
    DllCall("user32\SendMessageW", "Ptr", btn.Hwnd, "UInt", 0x00F7, "Ptr", 1, "Ptr", hIcon, "Ptr")

    ; 句柄需由本进程一直持有，直到按钮销毁，不能立即 DestroyIcon
    BakOpenHIcon := hIcon
    return true
}

; =====================================================================
; [系统级辅助函数：屏幕右下角免焦点悬浮卡片]
; =====================================================================
DestroyCurrentToast() {
    global toastGui, toastTitleCtrl, toastMsgCtrl, toastIconCtrl
    SetTimer(DestroyCurrentToast, 0)
    if (toastGui != "" && IsObject(toastGui)) {
        try {
            if (WinExist(toastGui.Hwnd))
                toastGui.Destroy()
        }
        toastGui := "", toastTitleCtrl := "", toastMsgCtrl := "", toastIconCtrl := ""
    }
}

ShowFloatingToast(title, msg, duration := 2000, isSuccess := true, opacity := 220) {
    global toastGui, toastTitleCtrl, toastMsgCtrl, toastIconCtrl
    bgColor := isSuccess ? "1E392A" : "5C2020"
    iconText := isSuccess ? "✔" : "✖"

    ; 每次都重新创建，确保文字改动后控件高度能同步刷新
    DestroyCurrentToast()

    toastGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")
    toastGui.BackColor := bgColor
    toastGui.MarginX := 20, toastGui.MarginY := 14

    toastGui.SetFont("s12 bold cFFFFFF", "Microsoft YaHei")
    toastIconCtrl := toastGui.Add("Text", "x20 y14 w24", iconText)
    toastTitleCtrl := toastGui.Add("Text", "x+2 yp w294", title)

    toastGui.SetFont("s10 norm cE0E0E0", "Microsoft YaHei")
    toastMsgCtrl := toastGui.Add("Text", "x20 y+8 w320", msg)

    ; 依据消息文字的实际高度自适应窗口高度，避免长路径被截断
    toastW := 360
    toastMsgCtrl.GetPos(&msgX, &msgY, &msgW, &msgH)
    toastH := msgY + msgH + 14
    posX := A_ScreenWidth - toastW - 25
    posY := A_ScreenHeight - toastH - 55

    toastGui.Show("x" posX " y" posY " w" toastW " h" toastH " NoActivate")
    try WinSetTransparent(opacity, toastGui.Hwnd)

    if (duration > 0)
        SetTimer(DestroyCurrentToast, -duration)
}

; =====================================================================
; [执行日志追加写入]
; 设计说明：
;   1. 日志固定保存在程序所在目录（编译后为 EXE 目录，未编译为脚本目录）；
;   2. 文件名固定为“公文.log”（v6.8.21 起；旧名“公文排版执行日志.txt”不再写入），
;      不附加日期和时间；
;   3. 每个文件处理完立即追加一行，成功、失败、成功但有提示都要记录；
;   4. 四个字段用 Tab 分隔：时间、结果、文件路径、说明；
;   5. 说明中的换行和 Tab 替换为空格，保证一个文件严格只占一行。
; =====================================================================

; 清理日志字段：把回车、换行、制表符替换为空格，并去掉首尾空白
SanitizeLogField(text) {
    s := String(text)
    s := StrReplace(s, "`r", " ")
    s := StrReplace(s, "`n", " ")
    s := StrReplace(s, "`t", " ")
    return Trim(s)
}

; 追加一条文件处理日志；写入成功返回 true，失败返回 false
AppendExecutionLog(result, filePath, detail := "") {
    ; v6.8.21：日志文件名改为“公文.log”。旧文件“公文排版执行日志.txt”保留但不再写入。
    logPath := A_ScriptDir "\公文.log"

    ; 先拼出前三列：时间、结果、文件路径
    line := FormatTime(, "yyyy-MM-dd HH:mm:ss")
          . "`t" SanitizeLogField(result)
          . "`t" SanitizeLogField(filePath)

    ; 只有存在说明时才追加第四列，避免成功且无提示的行尾出现多余 Tab
    cleanDetail := SanitizeLogField(detail)
    if (cleanDetail != "")
        line .= "`t" cleanDetail

    ; AHK 的 "a" 模式表示追加写入：文件不存在时自动创建，存在时写到末尾
    line .= "`r`n"
    try {
        f := FileOpen(logPath, "a", "UTF-8")
        f.Write(line)
        f.Close()
        return true
    } catch {
        return false
    }
}

; 打开执行日志：日志尚未生成时给出提示，已生成则交给系统默认程序打开
OpenLogFile(*) {
    ; v6.8.21：与 AppendExecutionLog 保持同一个文件名，否则“日志”按钮会打不开
    logPath := A_ScriptDir "\公文.log"
    if (!FileExist(logPath)) {
        ShowFloatingToast("日志尚未生成", "请先排版至少一个文件，程序会自动创建日志。", 3000, false)
        return
    }

    try {
        ; 用引号包住路径，兼容中文和带空格的目录；由系统默认程序打开 .txt
        Run('"' logPath '"')
    } catch {
        ShowFloatingToast("打开失败", "无法打开日志文件，请检查默认程序关联或文件权限。", 3000, false)
    }
}

MmToPt(mm) => Float(mm) * 2.83465

; =====================================================================
; [加载配置与注册表右键菜单管理]
; =====================================================================
; v6.8.26：旧版本默认把“页眉文字”写入注册表，里面可能包含单位名称、
; 处理对象等隐私内容。新版默认不记住页眉文字；这个函数在加载配置前运行，
; 对所有未勾选“记住文字”的用户方案和顶层配置把 HeaderText 清空为无内容。
; 只清这个字段，不动页眉字体/字号/横线等其它参数。
EraseUnkeptHeaderTextFromRegistry() {
    global REG_PROFILES_PATH, REG_PATH

    try {
        names := []
        Loop Reg, REG_PROFILES_PATH, "K" {
            n := A_LoopRegName
            if (n != "")
                names.Push(n)
        }
        for n in names {
            subKey := REG_PROFILES_PATH "\" n
            keep := 0
            try
                keep := Integer(RegRead(subKey, "HeaderKeep", 0))
            if (keep != 1)
                try
                    RegWrite("", "REG_SZ", subKey, "HeaderText")
        }
    }

    topKeep := 0
    try
        topKeep := Integer(GetConfigVal("Header", "HeaderKeep", 0))
    if (topKeep != 1)
        try
            RegWrite("", "REG_SZ", REG_PATH, "HeaderText")
}

LoadSettings() {
    ; v6.8.26：在读取顶层键之前先清理未勾选“记住文字”的隐私字段，
    ; 避免后续 ApplyProfileToGlobals 再把旧文字装载回运行时参数。
    EraseUnkeptHeaderTextFromRegistry()
    global MarginType := GetConfigVal("Margins", "MarginType", 1)
    global MarginTop := GetConfigVal("Margins", "MarginTop", 37)
    global MarginBottom := GetConfigVal("Margins", "MarginBottom", 35)
    global MarginLeft := GetConfigVal("Margins", "MarginLeft", 28)
    global MarginRight := GetConfigVal("Margins", "MarginRight", 26)

    global CustomTop := GetConfigVal("Margins", "CustomTop", 0.0)
    global CustomBottom := GetConfigVal("Margins", "CustomBottom", 0.0)
    global CustomLeft := GetConfigVal("Margins", "CustomLeft", 0.0)
    global CustomRight := GetConfigVal("Margins", "CustomRight", 0.0)

    global PageNumFormat := GetConfigVal("PageNumber", "PageNumFormat", "格式数字 — 1 — 居中")
    global SkipSinglePageNum := GetConfigVal("PageNumber", "SkipSinglePageNum", 1)
    ; v6.8.17 文档网格：默认值按边距预设推算（公文 → 22×28；其他 → 只指定行网格 44 行）
    global GridMode := Integer(GetConfigVal("Margins", "GridMode", (Integer(MarginType) == 1) ? 1 : 2))
    global GridLines := Integer(GetConfigVal("Margins", "GridLines", (Integer(MarginType) == 1) ? 22 : 44))
    global GridChars := Integer(GetConfigVal("Margins", "GridChars", (Integer(MarginType) == 1) ? 28 : 0))
    global CurrentFontScheme := GetConfigVal("FontScheme", "CurrentFontScheme", "党政机关公文格式")

    ; v6.8.20 页眉参数：与方案字段同名，切换方案时会由 ApplyProfileToGlobals() 覆盖
    global HeaderMode := Integer(GetConfigVal("Header", "HeaderMode", 0))
    global HeaderKeep := Integer(GetConfigVal("Header", "HeaderKeep", 0))
    ; v6.8.26 默认不记住页眉文字：即使旧注册表里留存 HeaderText，读取时也一律丢弃
    global HeaderText := HeaderKeep ? GetConfigVal("Header", "HeaderText", "") : ""
    global HeaderAlign := GetConfigVal("Header", "HeaderAlign", "居中")
    global HeaderMarginTop := GetConfigVal("Header", "HeaderMarginTop", 15)
    global HeaderFont := GetConfigVal("Header", "HeaderFont", "宋体")
    global HeaderSize := GetConfigVal("Header", "HeaderSize", "四号")
    global HeaderBold := Integer(GetConfigVal("Header", "HeaderBold", 0))
    global HeaderUnderline := Integer(GetConfigVal("Header", "HeaderUnderline", 0))

    global ExportPDF := GetConfigVal("Options", "ExportPDF", 0)
    global CreateBAK := GetConfigVal("Options", "CreateBAK", 1)
    global OptProtectTable := GetConfigVal("Options", "OptProtectTable", 0)
    global OptPreventOrphan := GetConfigVal("Options", "OptPreventOrphan", 0)
    global OptAttachmentNewPage := GetConfigVal("Options", "OptAttachmentNewPage", 1)
    global MaxHeadingLen := Integer(GetConfigVal("Options", "MaxHeadingLen", 35))
    global ContextActive := Integer(GetConfigVal("Options", "ContextActive", 0))

    if (MaxHeadingLen == 50) {
        MaxHeadingLen := 35
        SetConfigVal("Options", "MaxHeadingLen", 35)
    }

    global BakDir := GetConfigVal("Options", "BakDir", DefaultBakDir)
    if (BakDir == "")
        BakDir := DefaultBakDir

    LoadFontSchemes()
    ApplyFontScheme(CurrentFontScheme)
}

CreateTraySettingsMenu() {
    global SettingsMenu
    try A_TrayMenu.Delete()

    A_TrayMenu.Add("打开主窗口", (*) => ShowGUI())
    A_TrayMenu.Add("设置(&S)", SettingsMenu := Menu())
    SettingsMenu.Add("添加 Word 文件右键运行命令", (*) => RegisterDocumentContextMenu(true))
    SettingsMenu.Add("删除 Word 文件右键运行命令", (*) => RemoveDocumentContextMenu())
    A_TrayMenu.Add()
    A_TrayMenu.Add("退出(&X)", (*) => ExitApp())
}

GetContextMenuBaseClasses() {
    return [
        "HKCU\Software\Classes\.doc\shell",
        "HKCU\Software\Classes\.docx\shell",
        "HKCU\Software\Classes\SystemFileAssociations\.doc\Shell",
        "HKCU\Software\Classes\SystemFileAssociations\.docx\Shell",
        "HKCU\Software\Classes\Word.Document.12\shell",
        "HKCU\Software\Classes\Word.DocumentMacroEnabled.12\shell",
        "HKCU\Software\Classes\Word.Document.8\shell",
        "HKCU\Software\Classes\Word.DocumentMacroEnabled.8\shell"
    ]
}

RegisterDocumentContextMenu(showToast := true) {
    if (A_IsCompiled) {
        cmdPrefix := '"' . A_ScriptFullPath . '"', iconPath := A_ScriptFullPath
    } else if FileExist(A_ScriptDir "\公文自动排版.exe") {
        defaultExe := A_ScriptDir "\公文自动排版.exe"
        cmdPrefix := '"' . defaultExe . '"', iconPath := defaultExe
    } else {
        cmdPrefix := '"' . A_AhkPath . '" "' . A_ScriptFullPath . '"', iconPath := A_AhkPath
    }

    menuRunTitle     := "公文自动排版(&Z)"
    commandRun       := cmdPrefix . ' "%1"'

    hasError := false
    for _, basePath in GetContextMenuBaseClasses() {
        try {
            runKey := basePath "\DocFormatToolRun"
            RegCreateKey(runKey "\command")
            RegWrite(menuRunTitle, "REG_SZ", runKey)
            RegWrite(menuRunTitle, "REG_SZ", runKey, "MUIVerb")
            RegWrite(iconPath, "REG_SZ", runKey, "Icon")
            RegWrite(commandRun, "REG_SZ", runKey "\command")

            ; 清理旧版可能遗留的恢复菜单项
            try RegDeleteKey(basePath "\DocFormatToolRestore")
        } catch {
            hasError := true
        }
    }

    global ContextActive := 1
    SetConfigVal("Options", "ContextActive", 1)

    if (showToast) {
        if (!hasError) {
            ShowFloatingToast("右键命令添加成功", "已注册【" menuRunTitle "】命令", 2500, true)
        } else {
            ShowFloatingToast("注册提示", "部分项目写入受阻，请检查权限", 3500, false)
        }
    }
    return !hasError
}

RemoveDocumentContextMenu() {
    global ContextActive := 0
    SetConfigVal("Options", "ContextActive", 0)

    for _, basePath in GetContextMenuBaseClasses() {
        try RegDeleteKey(basePath "\DocFormatToolRun")
        try RegDeleteKey(basePath "\DocFormatToolRestore")
    }
    ShowFloatingToast("操作完成", "已删除 Word 文件右键菜单命令", 3000, true)
    return true
}

CheckAndRepairContextMenuPath() {
    regKeyPath := "HKCU\Software\Classes\.docx\shell\DocFormatToolRun\command"
    try {
        regCmd := RegRead(regKeyPath, "")
        regExePath := Trim(RegExReplace(regCmd, '"\s*"%1"$', ""), '" ')
        if (regExePath != "" && !FileExist(regExePath)) {
            RegisterDocumentContextMenu(false)
        }
    } catch {
        RegisterDocumentContextMenu(false)
    }
}

SaveAllCurrentSettings() {
    global ExportPDF, CreateBAK, BakDir, ActiveProfile, CurrentFontScheme

    ; 全局设置不随方案变化
    SetConfigVal("Options", "ExportPDF", ExportPDF)
    SetConfigVal("Options", "CreateBAK", CreateBAK)
    SetConfigVal("Options", "BakDir", BakDir)

    ; 当前方案参数写入注册表顶层，排版逻辑与静默模式继续读取原有键
    PersistProfileKeysToRegistry()
    SetConfigVal("Profile", "Profile_Active", ActiveProfile)

    ; 把当前生效参数同步回当前用户方案
    if (!IsBuiltinProfile(ActiveProfile) && ProfileExists(ActiveProfile))
        SaveProfileFromGlobals(ActiveProfile)
}

; =====================================================================
; [主排版流水线控制 (模式 B：GUI 列表批量排版)]
; =====================================================================
SetGUIControlsLock(isEnabled) {
    global uiBtnRun, uiBtnClear, uiBtnKill, uiBtnExport, uiBtnLog
    global uiProfile, uiBtnEditProfile, uiBtnSaveAsProfile
    global uiPDF, uiBAK, uiBakPath, uiBakBtn, uiBakOpen
    global uiSettingsMenu, GuiBatchRunning

    ; v6.8.22：与按钮置灰同步记录“批量是否在跑”，供列表右键菜单判断
    GuiBatchRunning := !isEnabled

    for ctrl in [uiBtnRun, uiBtnClear, uiBtnExport, uiProfile, uiBtnSaveAsProfile,
                 uiPDF, uiBAK, uiBakPath, uiBakBtn, uiBakOpen] {
        ctrl.Enabled := isEnabled
    }
    uiBtnKill.Enabled := true
    ; 日志按钮始终可用：批量排版期间也允许随时打开日志查看进度
    uiBtnLog.Enabled := true

    if (isEnabled)
        UpdateProfileButtons()
    else
        uiBtnEditProfile.Enabled := false

    ; 批量排版期间禁用“设置”菜单中的操作项（菜单本身不支持整体禁用）
    if (IsSet(uiSettingsMenu) && IsObject(uiSettingsMenu)) {
        for itemName in ["添加 Word 文件右键运行命令", "删除 Word 文件右键运行命令", "导出当前所有设置 (注册表 .reg)"] {
            try {
                if (isEnabled)
                    uiSettingsMenu.Enable(itemName)
                else
                    uiSettingsMenu.Disable(itemName)
            }
        }
    }
}

RunGUIProcessing(*) {
    global uiLV, uiSB, CreateBAK
    count := uiLV.GetCount()
    if (count == 0) {
        uiSB.SetText(" 任务列表为空！请先将 Word 文件或文件夹拖入列表中。")
        return
    }

    ; 排版会直接修改原文件，备份开关关闭时先给出明确提醒；默认按钮为“是”（备份）
    if (!EnsureBackupConfirmed()) {
        uiSB.SetText(" 已取消本次排版：备份开关未开启。")
        return
    }

    SaveAllCurrentSettings()
    uiSB.SetText(" 参数已保存。正在初始化运行环境...")
    SetGUIControlsLock(false)
    
    app := GetWordApplication(&appName)
    if (!app) {
        uiSB.SetText(" 运行失败：未检测到可用的 WPS 或 Microsoft Word！")
        SetGUIControlsLock(true)
        return
    }

    success_list := [], fail_list := [], warn_list := []
    ; 统计日志写入失败次数：只影响界面提示，不影响排版结果
    log_fail_count := 0

    Loop count {
        row := A_Index
        filePath := uiLV.GetText(row, 4)
        fileName := uiLV.GetText(row, 1)

        uiLV.Modify(row, "Col2", "▶ 正在排版...")
        uiSB.SetText(" 正在排版 (" row "/" count "): " fileName)
        Sleep(-1)

        if (!IsAppAlive(app))
            app := GetWordApplication(&appName)

        err_msg := "", warn_msg := ""
        if (FormatWordDocument(filePath, app, &err_msg, &warn_msg)) {
            status_text := "✔ 成功"
            log_result := "成功"
            log_detail := ""
            if (warn_msg != "") {
                ; v6.8.23：列表“结果”列不再追加“（有提示）”，只保留“✔ 成功”；
                ; 提示明细仍然写进日志说明列，状态栏与右下角汇总照旧统计“有提示”。
                ; PDF 导出失败属于用户关心的硬信息，仍单独标出。
                if InStr(warn_msg, "PDF")
                    status_text .= "（PDF 失败）"
                ; 成功但有提示：结果列仍记为成功，具体提示写入说明列
                log_result := "成功（有提示）"
                log_detail := warn_msg
                warn_list.Push({path: filePath, err: warn_msg})
            }
            ; 每处理完一个文件立即追加日志，避免批量中途退出时丢失记录
            if (!AppendExecutionLog(log_result, filePath, log_detail))
                log_fail_count += 1
            uiLV.Modify(row, "Col2", status_text)
            success_list.Push(filePath)
        } else {
            uiLV.Modify(row, "Col2", "❌ 失败: " err_msg)
            ; 失败同样要写日志，失败原因放入说明列
            if (!AppendExecutionLog("失败", filePath, err_msg))
                log_fail_count += 1
            fail_list.Push({path: filePath, err: err_msg})
        }
        Sleep(-1)
    }

    if (IsSet(app) && app != "") {
        try {
            if (app.Documents.Count == 0)
                app.Quit(0)
        }
    }

    ; 日志固定追加到程序所在目录，成功和失败都写；这里只负责提示写入结果
    if (log_fail_count > 0)
        logHint := "日志写入失败，请检查程序所在目录写入权限"
    else
        logHint := "日志已追加到程序所在目录的 公文.log"

    if (fail_list.Length > 0 || warn_list.Length > 0) {
        pdf_warn_count := 0
        for w in warn_list {
            if InStr(w.err, "PDF")
                pdf_warn_count += 1
        }
        msg := "排版结束！成功: " success_list.Length " 个，失败: " fail_list.Length " 个"
        if (warn_list.Length > 0)
            msg .= "，有提示: " warn_list.Length " 个"
        if (pdf_warn_count > 0)
            msg .= "（PDF 失败 " pdf_warn_count " 个）"
        msg .= "。"
        uiSB.SetText(" " msg " " logHint "。")
        toast_title := (fail_list.Length > 0) ? "文件排版完毕 (有失败)" : "文件排版完毕 (有提示)"
        ShowFloatingToast(toast_title, msg "`n" logHint, 3500, false)
    } else {
        msg := "全部处理成功！共完成 " success_list.Length " 个文件。"
        uiSB.SetText(" " msg " " logHint "。")
        ShowFloatingToast("文件排版完毕", msg "`n" logHint, 2000, true)
    }

    SetGUIControlsLock(true)
}

; =====================================================================
; [主排版流水线控制 (模式 A：静默多选 / 外部右键菜单排版)]
; =====================================================================
RunSilentProcessing(target_files) {
    unique_map := Map(), clean_targets := []
    for f in target_files {
        if (!unique_map.Has(f) && f != "") {
            unique_map[f] := true
            clean_targets.Push(f)
        }
    }

    processing_files := []
    for path in clean_targets {
        if InStr(FileExist(path), "D") {
            Loop Files, path "\*.*", "F" {
                ext := Format("{:L}", A_LoopFileExt)
                if ((ext == "doc" || ext == "docx") && SubStr(A_LoopFileName, 1, 2) != "~$")
                    processing_files.Push(A_LoopFilePath)
            }
        } else {
            SplitPath(path, &fileName, , &fileExt)
            ext := Format("{:L}", fileExt)
            if ((ext == "doc" || ext == "docx") && SubStr(fileName, 1, 2) != "~$")
                processing_files.Push(path)
        }
    }

    total_count := processing_files.Length
    if (total_count == 0)
        ExitApp()

    ; 静默排版同样受备份开关保护：关闭时先询问（默认“是”）
    if (!EnsureBackupConfirmed())
        ExitApp()

    app := GetWordApplication(&appName)
    if (!app) {
        ShowFloatingToast("排版引擎错误", "未检测到可用的 WPS 或 Word！", 3500, false)
        Sleep(3500)
        ExitApp()
    }

    success_list := [], fail_list := [], warn_list := []
    ; 统计日志写入失败次数：只影响提示，不影响排版结果
    log_fail_count := 0

    for idx, filePath in processing_files {
        SplitPath(filePath, &fileName)
        ShowFloatingToast("正在排版 (" idx "/" total_count ")", fileName, 0, true)

        if (!IsAppAlive(app))
            app := GetWordApplication(&appName)

        err_msg := "", warn_msg := ""
        if (FormatWordDocument(filePath, app, &err_msg, &warn_msg)) {
            log_result := "成功"
            log_detail := ""
            if (warn_msg != "") {
                ; 成功但有提示：结果列仍记为成功，具体提示写入说明列
                log_result := "成功（有提示）"
                log_detail := warn_msg
                warn_list.Push({path: filePath, err: warn_msg})
            }
            ; 每处理完一个文件立即追加日志，保证成功文件也有记录
            if (!AppendExecutionLog(log_result, filePath, log_detail))
                log_fail_count += 1
            success_list.Push(filePath)
        } else {
            ; 失败同样要写日志，失败原因放入说明列
            if (!AppendExecutionLog("失败", filePath, err_msg))
                log_fail_count += 1
            fail_list.Push({path: filePath, err: err_msg})
        }
    }

    if (IsSet(app) && app != "") {
        try {
            if (app.Documents.Count == 0)
                app.Quit(0)
        }
    }

    ; 日志提示：与 GUI 模式保持一致，始终告知追加结果
    if (log_fail_count > 0)
        logHint := "日志写入失败，请检查程序所在目录写入权限"
    else
        logHint := "日志已追加到程序所在目录的 公文.log"

    if (fail_list.Length > 0 || warn_list.Length > 0) {
        pdf_warn_count := 0
        for w in warn_list {
            if InStr(w.err, "PDF")
                pdf_warn_count += 1
        }
        msg := "成功: " success_list.Length " 个，失败: " fail_list.Length " 个"
        if (warn_list.Length > 0)
            msg .= "，有提示: " warn_list.Length " 个"
        if (pdf_warn_count > 0)
            msg .= "（PDF 失败 " pdf_warn_count " 个）"
        msg .= "`n" logHint
        toast_title := (fail_list.Length > 0) ? "文件排版完毕 (有失败)" : "文件排版完毕 (有提示)"
        ShowFloatingToast(toast_title, msg, 3500, false)
        Sleep(3500)
    } else {
        msg := "已顺利完成 " success_list.Length " 个文件的自动排版。"
        ShowFloatingToast("文件自动排版成功", msg "`n" logHint, 2000, true)
        Sleep(2000)
    }
    ExitApp()
}

; =====================================================================
; [底层排版辅助函数区]
; =====================================================================

GetVisualLength(text) {
    len := 0.0
    Loop Parse text
        len += (Ord(A_LoopField) <= 127) ? 0.5 : 1.0
    return len
}

IsAppAlive(app) {
    try {
        test := app.Version
        return true
    } catch {
        return false
    }
}

GetWordApplication(&appName) {
    global ManagedWordPIDs
    candidates := [
        ["Kwps.Application", "WPS 文字 (Kwps)"],
        ["WPS.Application", "WPS 文字 (WPS)"],
        ["Word.Application", "Microsoft Word"]
    ]
    for item in candidates {
        try {
            app := ComObject(item[1])
            app.Visible := false
            try app.DisplayAlerts := 0
            appName := item[2]
            
            try {
                prev_Detect := A_DetectHiddenWindows
                DetectHiddenWindows(true)
                original_caption := app.Caption
                unique_caption := "DocFormatTool_Instance_" A_TickCount
                app.Caption := unique_caption
                if (hw := WinExist(unique_caption)) {
                    if (pid := WinGetPID("ahk_id " hw))
                        ManagedWordPIDs[pid] := true
                }
                app.Caption := original_caption
                DetectHiddenWindows(prev_Detect)
            }
            return app
        }
    }
    appName := ""
    return ""
}

SetRangeFont(rng, fName, fSize, fBold) {
    f := rng.Font
    f.NameFarEast := fName
    f.NameAscii   := fName
    f.Size        := fSize
    f.Bold        := fBold
}

; v6.8.28：安全地给段落应用 Word 内置段落样式（正文/标题样式化）。
; styleId 使用 WdBuiltinStyle 整数常量，Word 与 WPS 通常都支持：
;   -1=wdStyleNormal(正文), -63=wdStyleTitle(标题), -75=wdStyleSubtitle(副标题),
;   -2..-5=wdStyleHeading1..4(标题1-4)。设置样式可能重置直接格式，
;   调用方应在本函数之后重新应用本程序需要的字体/段落格式。
; 失败时静默忽略，避免 WPS 个别版本或受保护文档中断排版。
TrySetParaStyle(para, styleId) {
    if (!IsObject(para) || !IsInteger(styleId) || styleId == 0)
        return false
    try {
        para.Range.Style := styleId
        return true
    }
    try {
        para.Style := styleId
        return true
    }
    return false
}

; v6.8.29：在段落应用样式之前，先把文档里的内置段落样式（正文/标题/副标题/
; 标题1~4）的段前、段后、行前、行后间距以及行距调整为国标要求（无段前段后、
; 方案指定的固定行距）。这样后续给段落应用样式时，段落会继承已经调整好的
; 间距，不会带上 Word/WPS 样式自带的 8 pt 段后距、单倍行距等。
; v6.8.31：在段落应用样式之前，先把文档里的内置段落样式（正文/标题/副标题/
; 标题1~4）的段前、段后、行前、行后间距、行距以及字体字号加粗调整为国标要求。
; 这样后续给段落应用样式时，段落会继承已经调整好的样式字体格式，不会带上
; Word/WPS 样式自带的 8 pt 段后距、单倍行距或错误字体。
AdjustBuiltinStyleSpacing(doc) {
    global LINE_SPACING_RULE, LINE_SPACING_BODY, ActiveStyle

    ; v6.8.32：不再依赖 FormatWordDocument() 里可能尚未赋值的局部 ptMTitle/ptH1/... 等
    ; 全局变量，而是在函数内部根据 ActiveStyle 重新用 ParseFontSize() 计算各级字号，
    ; 避免“This global variable has not been assigned a value”错误。
    ptMTitle   := ParseFontSize(ActiveStyle.MTitleSize, 22.0)
    ptSubTitle := ParseFontSize(ActiveStyle.SubTitleSize, 16.0)
    ptH1       := ParseFontSize(ActiveStyle.H1Size, 16.0)
    ptH2       := ParseFontSize(ActiveStyle.H2Size, 16.0)
    ptH3       := ParseFontSize(ActiveStyle.H3Size, 16.0)
    ptH4       := ParseFontSize(ActiveStyle.H4Size, 16.0)
    ptBody     := ParseFontSize(ActiveStyle.BodySize, 16.0)

    ; WdBuiltinStyle 常量：-1=正文(Normal), -63=标题(Title), -75=副标题(Subtitle),
    ; -2~-5=标题1~4(Heading1~4)
    ; 每个样式对应的方案字体/字号/加粗/首行缩进，先调整样式本身的格式，再应用到段落。
    ; 正文与标题1~4全部首行缩进2字符，套用样式后缩进不会丢失。
    styleMap := [
        {id: -1,  font: ActiveStyle.BodyFont,    size: ptBody,    bold: ActiveStyle.BodyBold,    firstLineIndent: 2},
        {id: -63, font: ActiveStyle.MTitleFont,  size: ptMTitle,  bold: ActiveStyle.MTitleBold,  firstLineIndent: 2},
        {id: -75, font: ActiveStyle.SubTitleFont,size: ptSubTitle,bold: ActiveStyle.SubTitleBold,firstLineIndent: 0},
        {id: -2,  font: ActiveStyle.H1Font,      size: ptH1,      bold: ActiveStyle.H1Bold,      firstLineIndent: 2},
        {id: -3,  font: ActiveStyle.H2Font,      size: ptH2,      bold: ActiveStyle.H2Bold,      firstLineIndent: 2},
        {id: -4,  font: ActiveStyle.H3Font,      size: ptH3,      bold: ActiveStyle.H3Bold,      firstLineIndent: 2},
        {id: -5,  font: ActiveStyle.H4Font,      size: ptH4,      bold: ActiveStyle.H4Bold,      firstLineIndent: 2}
    ]
    for each, item in styleMap {
        try {
            st := doc.Styles.Item(item.id)
            ; 调整段间距与首行缩进（先调整样式格式，再应用到段落）
            pf := st.ParagraphFormat
            pf.SpaceBeforeAuto := 0
            pf.SpaceAfterAuto  := 0
            pf.LineUnitBefore  := 0
            pf.LineUnitAfter   := 0
            pf.SpaceBefore     := 0
            pf.SpaceAfter      := 0
            pf.LineSpacingRule := LINE_SPACING_RULE
            pf.LineSpacing     := LINE_SPACING_BODY
            pf.CharacterUnitFirstLineIndent := item.firstLineIndent   ; 样式本身带2字符首行缩进
            pf.FirstLineIndent := 0                                   ; 按字符单位，避免磅值混淆
            ; 调整样式本身的字体字号加粗
            f := st.Font
            f.NameFarEast := item.font
            f.NameAscii   := item.font
            f.Name        := item.font
            f.Size        := item.size
            f.Bold        := item.bold
        }
        ; 个别环境下 WPS 可能不支持某个内置样式，静默跳过，不影响排版
    }
}

; v6.8.30：在文档中查找或创建名为“主送机关”的自定义段落样式。该样式本身即为顶格：
;   Alignment:=0（左对齐/顶格），CharacterUnitFirstLineIndent:=0、FirstLineIndent:=0（无首行缩进），
;   SpaceBefore/SpaceAfter:=0（无段前段后），行距沿用方案的全文行距。
; 如果文档中已存在同名样式，直接返回；不存在则尝试用 doc.Styles.Add 创建（WdStyleType=1 表示段落样式）。
; 创建/获取失败时返回空字符串，调用方自行降级为直接格式设置。
GetOrCreateMainAddresseeStyle(doc) {
    global LINE_SPACING_RULE, LINE_SPACING_BODY
    styleName := "主送机关"

    ; 先尝试获取已有样式。已存在时也要把首行缩进重置为 0，避免旧文档里残留的
    ; 缩进值让主送机关行误加首行缩进。
    try {
        st := doc.Styles.Item(styleName)
        if (IsObject(st)) {
            pf := st.ParagraphFormat
            pf.Alignment := 0
            pf.CharacterUnitFirstLineIndent := 0
            pf.FirstLineIndent := 0
            pf.SpaceBeforeAuto := 0
            pf.SpaceAfterAuto  := 0
            pf.LineUnitBefore  := 0
            pf.LineUnitAfter   := 0
            pf.SpaceBefore     := 0
            pf.SpaceAfter      := 0
            return styleName
        }
    }

    ; 不存在则尝试创建：doc.Styles.Add(Name, StyleType)，StyleType=1 为段落样式
    try {
        st := doc.Styles.Add(styleName, 1)
        pf := st.ParagraphFormat
        pf.Alignment := 0
        pf.CharacterUnitFirstLineIndent := 0
        pf.FirstLineIndent := 0
        pf.SpaceBeforeAuto := 0
        pf.SpaceAfterAuto  := 0
        pf.LineUnitBefore  := 0
        pf.LineUnitAfter   := 0
        pf.SpaceBefore     := 0
        pf.SpaceAfter      := 0
        pf.LineSpacingRule := LINE_SPACING_RULE
        pf.LineSpacing     := LINE_SPACING_BODY
        return styleName
    }
    ; WPS/旧模板/受保护文档可能不支持 Styles.Add，静默返回空，由调用方降级
    return ""
}

SetParagraphText(paragraph, newText) {
    r := paragraph.Range.Duplicate
    if (r.End > r.Start)
        r.End := r.End - 1   ; 排除段落标记，只改文字部分

    ; v6.8.13：保留原文里的分页符/分节符（Word 统一存为 Chr(12)）。
    ; 直接 r.Text := newText 会把段末的分页符一起覆盖掉，导致页数变少、分节被合并。
    origText := r.Text
    breakSuffix := ""
    if (InStr(origText, Chr(12)) && !InStr(newText, Chr(12))) {
        pos := 1
        while (pos := InStr(origText, Chr(12), , pos)) {
            breakSuffix .= Chr(12)
            pos += 1
        }
    }
    r.Text := newText . breakSuffix
}

; ---------------------------------------------------------------------
; [分页符/分节符保护（v6.8.13）]
; Word 把“手动分页符”和各类“分节符”都存成段落文字里的 Chr(12)，
; 段落也可能单独设置“段前分页”（PageBreakBefore）。这类段落看起来是空段，
; 但删掉就会丢页、并掉分节，所以删除空段、改写文字前必须先判断。
; ---------------------------------------------------------------------

; 文字里是否含分页符/分节符
HasPageBreakChar(text) {
    return InStr(text, Chr(12)) > 0
}

; 段落是否受保护：含分页符/分节符，或设置了“段前分页”
IsPageBreakProtected(paragraph) {
    if (!IsObject(paragraph))
        return false
    try {
        if (HasPageBreakChar(paragraph.Range.Text))
            return true
    }
    try {
        if (paragraph.Format.PageBreakBefore)
            return true
    }
    return false
}

; 段落是否“只有分页符、没有文字”（去掉 Chr(12) 和空白后为空）
IsBreakOnlyParagraph(paragraph) {
    if (!IsObject(paragraph))
        return false
    try {
        raw := RTrim(paragraph.Range.Text, "`r`n")
        if (!HasPageBreakChar(raw))
            return false
        return TrimAllSpaces(StrReplace(raw, Chr(12), "")) == ""
    }
    return false
}

; 让“附件1”标记段另起一页（附件另面编排，GB/T 9704-2012 7.3.8）。
; v6.8.15：改为插入分节符（下一页）将附件划分到新节，断开新节页脚链接并
; 将页码重置为从 1 开始（附件页码单独编号）。若标记本身已在新页或已设置
; 段前分页/已有分页符，则不再重复处理，避免多出空白页。
ApplyAttachmentPageBreak(doc, p_mark, mark_idx) {
    global wdSectionBreakNextPage, gAttachSectionIndex
    if (!IsObject(p_mark) || mark_idx <= 1)
        return

    ; 判断附件是否已经在新页：只有“段前分页”已设，或“上一段末尾”有分页符/分节符，
    ; 才说明附件标记段本身已经位于页首。若“附件标记”段自己含 Chr(12)，通常表示
    ; 页码在附件文字之后才换页，附件标记本身还不在新页，不能据此跳过处理。
    alreadyNewPage := false
    try {
        if (p_mark.Format.PageBreakBefore)
            alreadyNewPage := true
    }
    try {
        if (!alreadyNewPage && HasPageBreakChar(doc.Paragraphs(mark_idx - 1).Range.Text))
            alreadyNewPage := true
    }

    ; 若附件尚未在新页，在“附件标记”段落最前面插入分节符（下一页），新建一节。
    ; 插入分节符可让后续页码独立编号；若插入失败，回退到段前分页（只换页，页码仍顺延）。
    if (!alreadyNewPage) {
        try {
            r := p_mark.Range.Duplicate
            r.Collapse(1)
            r.InsertBreak(wdSectionBreakNextPage)
        } catch {
            try p_mark.Format.PageBreakBefore := true
        }
    }

    ; 分节符插入后，重新定位“附件标记”段落所在的节，用于页码重置。
    ; 注意：不能直接用 p_mark.Range.Paragraphs(1)。InsertBreak 之后 p_mark 的
    ; 范围可能仍指向分节符之前的旧段落，Sections(1) 就会拿到旧节，
    ; gAttachSectionIndex 记错，页码重置落空。
    ; 这里先在 mark_idx 附近按“附件N”文本重新查找，最可靠；找不到再退回索引。
    p_mark2 := ""
    try {
        for off in [0, 1, -1, 2, -2] {
            idx2 := mark_idx + off
            if (idx2 < 1 || idx2 > doc.Paragraphs.Count)
                continue
            cand := doc.Paragraphs(idx2)
            candText := TrimAllSpaces(StrReplace(RTrim(cand.Range.Text, "`r`n"), Chr(12), ""))
            if (RegExMatch(candText, "^附件\s*[0-9０-９一二三四五六七八九十]+\s*[:：]?\s*$")) {
                p_mark2 := cand
                break
            }
        }
    }
    if (!IsObject(p_mark2)) {
        try p_mark2 := doc.Paragraphs(mark_idx)
        catch
            p_mark2 := p_mark
    }

    sec := ""
    try sec := p_mark2.Range.Sections(1)
    if (!IsObject(sec)) {
        try {
            cnt := doc.Sections.Count
            if (cnt >= 1)
                sec := doc.Sections(cnt)
        }
    }

    if (IsObject(sec)) {
        try gAttachSectionIndex := sec.Index
        ps := ""
        try ps := sec.PageSetup
        if (IsObject(ps)) {
            ; 断开该节页脚与上一节的链接（首页、奇偶页、普通页脚）
            try sec.Footers(1).LinkToPrevious := false
            try sec.Footers(1).Range.Text := ""
            try {
                if (ps.DifferentFirstPageHeaderFooter) {
                    try sec.Footers(2).LinkToPrevious := false
                    try sec.Footers(2).Range.Text := ""
                }
            }
            try {
                if (ps.OddAndEvenPagesHeaderFooter) {
                    try sec.Footers(3).LinkToPrevious := false
                    try sec.Footers(3).Range.Text := ""
                }
            }
            ; 强制设置：本节页码从 1 开始。
            ; 重要（WPS 兼容）：WPS 的 PageSetup 对象根本没有 PageNumbering 属性，
            ; 所以必须通过 Footers(x).PageNumbers 设置，WPS 才认。
            ; Word 两者都支持，这里先写 Footers.PageNumbers，再补 PageSetup 兜底。
            try {
                pnums := sec.Footers(1).PageNumbers
                pnums.RestartNumberingAtSection := true
                pnums.StartingNumber := 1
            }
            try {
                pn := ps.PageNumbering
                pn.RestartNumberingAtSection := true
                pn.StartingNumber := 1
            }
            ; 部分 Word 版本也需要在页眉页脚集合上确保不链接
            try sec.Headers(1).LinkToPrevious := false
            try {
                if (ps.DifferentFirstPageHeaderFooter)
                    sec.Headers(2).LinkToPrevious := false
            }
            try {
                if (ps.OddAndEvenPagesHeaderFooter)
                    sec.Headers(3).LinkToPrevious := false
            }
        }
    }

    try {
        if (IsObject(p_mark2))
            try p_mark2.Format.PageBreakBefore := false
    }
    try {
        if (IsObject(p_mark))
            try p_mark.Format.PageBreakBefore := false
    }
}

TrimAllSpaces(str) {
    return RegExReplace(str, "^[\s\x{00A0}\x{2000}-\x{200D}\x{FEFF}\x{3000}]+|[\s\x{00A0}\x{2000}-\x{200D}\x{FEFF}\x{3000}]+$", "")
}

LTrimAllSpaces(str) {
    return RegExReplace(str, "^[\s\x{00A0}\x{2000}-\x{200D}\x{FEFF}\x{3000}]+", "")
}

IsEndWithPunct(text) {
    clean_text := TrimAllSpaces(text)
    return (clean_text != "") && (InStr("。！？；：，、!?;:,", SubStr(clean_text, -1)) > 0)
}

IsEndWithColon(text) {
    clean_text := TrimAllSpaces(text)
    return (clean_text != "") && (InStr("：:", SubStr(clean_text, -1)) > 0)
}

; 判断一段文字是否采用一至四级标题的编号样式。
; 标题区扫描时用它决定“副标题块”到哪里结束，因此正则必须与正文排版处
; （FormatWordDocument 中的 m_h1～m_h4）保持一致，避免两处判断结果不同。
IsHeadingText(text) {
    t := TrimAllSpaces(text)
    if (t == "")
        return false
    ; 一级：“一、”“二.”等；排除“二〇二六年”这类年份写法
    if (RegExMatch(t, "^([一二三四五六七八九十百]+)\s*([、.．·])\s*(.*)$", &m1) && !RegExMatch(m1[3], "^[一二三四五六七八九十百]+[、年]"))
        return true
    ; 二级：“（一）”“(一)”
    if (RegExMatch(t, "^(\x28|（)\s*([一二三四五六七八九十百]+)\s*(\x29|）)\s*(.*)$"))
        return true
    ; 三级：“1.”“1、”“1．”
    if (RegExMatch(t, "^(\d+)\s*([.．、·])(?!\d)\s*(.*)$"))
        return true
    ; 四级：“（1）”“(1)”
    if (RegExMatch(t, "^(\x28|（)\s*(\d+)\s*(\x29|）)\s*(.*)$"))
        return true
    return false
}

; 判断一段文字是否为“2026年10月26日”这类成文日期。
; 标题区扫描时日期要作为副标题，不能被三级标题正则误伤，也不能被
; 文末落款逻辑误认为标题，所以单独抽一个判断函数。
IsDateText(text) {
    return RegExMatch(TrimAllSpaces(text), "^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日") > 0
}

GetCleanText(paragraph) {
    raw := RTrim(paragraph.Range.Text, "`r`n")
    cleaned := Trim(raw, " `t`v" . Chr(12288) . Chr(160))
    cleaned := RegExReplace(cleaned, "^[\x{2000}-\x{200D}\x{FEFF}]+|[\x{2000}-\x{200D}\x{FEFF}]+$", "")
    if (raw != cleaned && cleaned != "") {
        try {
            if (paragraph.Range.InlineShapes.Count == 0 && paragraph.Range.Fields.Count == 0) {
                SetParagraphText(paragraph, cleaned)
            } else {
                leadLen := StrLen(raw) - StrLen(LTrim(raw, " `t`v" . Chr(12288) . Chr(160)))
                if (leadLen > 0) {
                    rLead := paragraph.Range.Duplicate
                    rLead.SetRange(rLead.Start, rLead.Start + leadLen)
                    rLead.Delete()
                }
            }
        }
    }
    return cleaned
}

CleanAllBlankParagraphs(doc) {
    count := doc.Paragraphs.Count
    Loop count {
        if (Mod(A_Index, 20) == 0)
            Sleep(-1)
        if (doc.Paragraphs.Count <= 1)
            break
        i := count - A_Index + 1
        p := doc.Paragraphs(i)
        try {
            if (p.Range.Tables.Count > 0 || p.Range.InlineShapes.Count > 0)
                continue
        }
        txt := TrimAllSpaces(RTrim(p.Range.Text, "`r`n"))
        if (txt == "") {
            ; 只含分页符/分节符的段落、或设置了段前分页的段落不是空行，
            ; 删掉会丢页、并掉分节（v6.8.13）
            if (IsPageBreakProtected(p))
                continue
            try p.Range.Delete()
        }
    }
}

RemoveTrailingEmptyParagraphs(doc) {
    while (doc.Paragraphs.Count > 1) {
        last_p := doc.Paragraphs(doc.Paragraphs.Count)
        try {
            if (last_p.Range.Tables.Count > 0 || last_p.Range.InlineShapes.Count > 0)
                break
        }
        txt := TrimAllSpaces(RTrim(last_p.Range.Text, "`r`n"))
        if (txt == "") {
            ; 文末只剩分页符的段落要保留，否则文档末尾的空白页会消失（v6.8.13）
            if (IsPageBreakProtected(last_p))
                break
            try last_p.Range.Delete()
            catch
                break
        } else {
            break
        }
    }
}

; ---------------------------------------------------------------------
; [表格保护：记录并恢复表格原有字符/段落格式]
; 勾选“表格保护与字号适配”时，排版过程中的正文级格式会临时作用于表格，
; 这里在排版前保存表格内每个段落的字体与段落格式，排版完成后原样写回。
; ---------------------------------------------------------------------
CaptureTableFormats(doc) {
    snaps := []
    for tbl in doc.Tables {
        cellsSnap := []
        for cell in tbl.Range.Cells {
            parasSnap := []
            for p in cell.Range.Paragraphs {
                f := p.Range.Font
                pf := p.Format
                parasSnap.Push({
                    fontFarEast: f.NameFarEast, fontAscii: f.NameAscii,
                    size: f.Size, bold: f.Bold,
                    align: pf.Alignment,
                    charFirst: pf.CharacterUnitFirstLineIndent,
                    firstIndent: pf.FirstLineIndent,
                    charLeft: pf.CharacterUnitLeftIndent,
                    leftIndent: pf.LeftIndent,
                    charRight: pf.CharacterUnitRightIndent,
                    rightIndent: pf.RightIndent,
                    lineRule: pf.LineSpacingRule,
                    lineSpacing: pf.LineSpacing,
                    spaceBefore: pf.SpaceBefore,
                    spaceAfter: pf.SpaceAfter,
                    keepNext: pf.KeepWithNext
                })
            }
            cellsSnap.Push(parasSnap)
        }
        snaps.Push(cellsSnap)
    }
    return snaps
}

; 写回单个属性；wdUndefined(9999999) 表示原格式不统一，跳过以免报错
RestoreRangeProp(obj, prop, val) {
    if ((val = "") || (val = 9999999))
        return
    try obj.%prop% := val
}

RestoreTableFormats(doc, snaps) {
    if (!IsObject(snaps))
        return
    tIdx := 0
    for tbl in doc.Tables {
        tIdx += 1
        if (tIdx > snaps.Length)
            break
        cellsSnap := snaps[tIdx]
        cIdx := 0
        for cell in tbl.Range.Cells {
            cIdx += 1
            if (cIdx > cellsSnap.Length)
                break
            parasSnap := cellsSnap[cIdx]
            pIdx := 0
            for p in cell.Range.Paragraphs {
                pIdx += 1
                if (pIdx > parasSnap.Length)
                    break
                ps := parasSnap[pIdx]

                f := ""
                try {
                    f := p.Range.Font
                    f.NameFarEast := ps.fontFarEast
                    f.NameAscii := ps.fontAscii
                }
                if (IsObject(f)) {
                    RestoreRangeProp(f, "Size", ps.size)
                    RestoreRangeProp(f, "Bold", ps.bold)
                }

                try pf := p.Format
                catch
                    continue
                RestoreRangeProp(pf, "Alignment", ps.align)
                RestoreRangeProp(pf, "CharacterUnitFirstLineIndent", ps.charFirst)
                RestoreRangeProp(pf, "FirstLineIndent", ps.firstIndent)
                RestoreRangeProp(pf, "CharacterUnitLeftIndent", ps.charLeft)
                RestoreRangeProp(pf, "LeftIndent", ps.leftIndent)
                RestoreRangeProp(pf, "CharacterUnitRightIndent", ps.charRight)
                RestoreRangeProp(pf, "RightIndent", ps.rightIndent)
                RestoreRangeProp(pf, "LineSpacingRule", ps.lineRule)
                RestoreRangeProp(pf, "LineSpacing", ps.lineSpacing)
                RestoreRangeProp(pf, "SpaceBefore", ps.spaceBefore)
                RestoreRangeProp(pf, "SpaceAfter", ps.spaceAfter)
                RestoreRangeProp(pf, "KeepWithNext", ps.keepNext)
            }
        }
    }
}

; ---------------------------------------------------------------------
; [页码安全写入辅助函数]
; 说明：只替换本程序书签标记的页码块；页脚中的其他文字、表格、图片
;       和用户自行插入的页码均保留，无法安全处理时记录提示。
; ---------------------------------------------------------------------
GetFooterVisibleText(obj) {
    txt := obj.Range.Text
    txt := StrReplace(txt, Chr(13), "")
    txt := StrReplace(txt, Chr(10), "")
    txt := StrReplace(txt, Chr(12), "")
    return TrimAllSpaces(txt)
}

GetFooterPageFieldInfo(footer) {
    pageCount := 0, numPagesCount := 0, otherCount := 0
    fields := footer.Range.Fields
    Loop fields.Count {
        f := fields.Item(A_Index)
        if (f.Type == 33)
            pageCount += 1
        else if (f.Type == 26)
            numPagesCount += 1
        else
            otherCount += 1
    }
    return {page: pageCount, numPages: numPagesCount, other: otherCount}
}

IsFooterOnlyPageNumber(footer) {
    info := GetFooterPageFieldInfo(footer)
    if (info.other > 0 || info.page == 0 || info.page > 1 || info.numPages > 1)
        return false

    ; 页脚含表格、图片、内容控件、超链接、书签、批注或修订时，绝不整段清空
    if (footer.Range.Tables.Count > 0)
        return false
    if (footer.Range.InlineShapes.Count > 0)
        return false
    if (footer.Range.Bookmarks.Count > 0)
        return false
    if (footer.Range.Hyperlinks.Count > 0)
        return false
    try {
        if (footer.Range.ShapeRange.Count > 0)
            return false
    }
    try {
        if (footer.Range.ContentControls.Count > 0)
            return false
    }
    try {
        if (footer.Range.Comments.Count > 0)
            return false
    }
    try {
        if (footer.Range.Revisions.Count > 0)
            return false
    }

    txt := GetFooterVisibleText(footer)
    ; 兼容本程序生成的“— 1 —”“1 / 5”“1”以及常见的“第 1 页 共 5 页”；
    ; 字段结果暂时为空时也允许通过，避免未更新域时误判。
    if RegExMatch(txt, "^\s*[—–\-－]?\s*\d*\s*[—–\-－]?\s*$")
        return true
    if RegExMatch(txt, "^\s*\d*\s*[/／]\s*\d*\s*$")
        return true
    if RegExMatch(txt, "^\s*第\s*\d*\s*页\s*(共\s*\d*\s*页)?\s*$")
        return true
    return false
}

FooterHasPageField(footer) {
    fields := footer.Range.Fields
    Loop fields.Count {
        f := fields.Item(A_Index)
        if (f.Type == 33 || f.Type == 26)
            return true
    }
    return false
}

UpdateFooterPageFields(footer) {
    fields := footer.Range.Fields
    Loop fields.Count {
        f := fields.Item(A_Index)
        if (f.Type == 33 || f.Type == 26)
            try f.Update()
    }
}

FooterTypeLabel(type) {
    if (type == 2)
        return "首页页脚"
    if (type == 3)
        return "偶数页页脚"
    return "页脚"
}

RemoveFooterPageBlock(footer) {
    ; 清理本程序写入的页码书签块。除精确书签名外，也清理因取消
    ; “链接到前一节”而被复制的书签，避免重复页码。
    bms := footer.Range.Bookmarks
    removed := false
    guard := 0
    i := 1
    while (i <= bms.Count) {
        guard += 1
        if (guard > 100)
            throw Error("页脚页码书签清理失败")
        name := bms.Item(i).Name
        if (SubStr(name, 1, 10) == "DocFmt_PN_") {
            bm := bms.Item(i)
            try bm.Range.Delete()
            try bm.Delete()
            removed := true
            i := 1
            continue
        }
        i += 1
    }
    return removed
}

WriteFooterPageBlock(footer, paraIndex, formatStr, bmName) {
    global ActiveStyle
    ; v6.8.14：页码字体字号改为读取当前方案（默认 宋体 四号 不加粗），
    ; 由编辑窗口“字体字号”表格最后一行“页码”控制，不再写死。
    fontPageNum := ActiveStyle.PageNumFont
    sizePageNum := ParseFontSize(ActiveStyle.PageNumSize, 14.0)
    boldPageNum := ActiveStyle.PageNumBold
    alignVal := InStr(formatStr, "居左") ? 0 : (InStr(formatStr, "居右") ? 2 : 1)
    wdFieldPage := 33, wdFieldNumPages := 26

    target := footer.Range.Paragraphs.Item(paraIndex).Range
    r := target.Duplicate
    r.Collapse(1)

    if (InStr(formatStr, "格式数字")) {
        r.Text := " —"
        r.Collapse(1)
        footer.Range.Fields.Add(r, wdFieldPage, , false)
        rStart := target.Duplicate
        rStart.Collapse(1)
        rStart.Text := "— "
    } else if (InStr(formatStr, "加总页码")) {
        r.Text := " / "
        r.Collapse(0)
        footer.Range.Fields.Add(r, wdFieldNumPages, , false)
        rStart := target.Duplicate
        rStart.Collapse(1)
        footer.Range.Fields.Add(rStart, wdFieldPage, , false)
    } else {
        footer.Range.Fields.Add(r, wdFieldPage, , false)
    }
    ; v6.8.15：添加页码字段后显式更新，使 PAGE 字段使用新节的起始页码（从 1 开始）
    try footer.Range.Fields.Update()

    ; 重新获取段落范围，确保包含刚插入的域和文字
    target := footer.Range.Paragraphs.Item(paraIndex).Range
    target.ParagraphFormat.Alignment := alignVal
    target.ParagraphFormat.CharacterUnitFirstLineIndent := 0
    target.ParagraphFormat.FirstLineIndent := 0
    target.ParagraphFormat.CharacterUnitLeftIndent := 0
    target.ParagraphFormat.LeftIndent := 0
    target.ParagraphFormat.CharacterUnitRightIndent := 0
    target.ParagraphFormat.RightIndent := 0
    target.ParagraphFormat.LineSpacingRule := 0 ; 0 = 单倍行距
    target.ParagraphFormat.SpaceBefore := 0
    target.ParagraphFormat.SpaceAfter := 0
    SetRangeFont(target, fontPageNum, sizePageNum, boldPageNum)

    if (footer.Range.Bookmarks.Exists(bmName))
        footer.Range.Bookmarks.Item(bmName).Delete()
    footer.Range.Bookmarks.Add(bmName, target)
}

; =====================================================================
; [页眉写入（v6.8.21）]
; 设计说明：
;   1. 三态由方案 HeaderMode 决定：0 = 不改变、1 = 添加·替换、2 = 清空；
;   2. 只写主页眉 Headers(1)。只有该节本来已启用“首页不同”/“奇偶页不同”时，
;      才顺带处理 Headers(2)（首页页眉）/ Headers(3)（偶数页页眉）；不主动打开
;      这两个版式开关——打开会改变整篇文档的版面，属于用户没有要求的副作用；
;   3. 写之前必须先 LinkToPrevious := false：页眉默认“链接到前一节”，
;      不断开的话 Headers(x) 拿到的是上一节的同一份内容，会改到别的节；
;   4. 页眉文字留空 = 用文档文件名（不含扩展名），与 PDF 加页码工具语义一致；
;   5. 任何一节失败都只记提示、不抛异常：页眉不是排版必需项，
;      失败不能把整个文件判成“排版失败”。
; 参数：doc = Word/WPS 的 Document COM 对象
; 返回：提示串（空串表示没有异常/没有覆盖提示）
; =====================================================================
SetDocumentHeaders(doc) {
    global HeaderMode, HeaderText, HeaderAlign, HeaderMarginTop
    global HeaderFont, HeaderSize, HeaderBold, HeaderUnderline
    global ActiveStyle

    mode := Integer(HeaderMode)
    ; “不改变”（以及任何异常值）：直接返回，一个节都不碰，原有页眉原样保留
    if (mode != 1 && mode != 2)
        return ""

    ; 页眉文字：留空是合法值，含义是“用文档文件名”。doc.Name 带扩展名，
    ; 这里用 SplitPath 去掉扩展名，只留主文件名。
    headerText := Trim(HeaderText)
    if (mode == 1 && headerText == "") {
        docNameNoExt := ""
        try SplitPath(doc.Name, , , , &docNameNoExt)
        headerText := docNameNoExt
    }

    ; 对齐：Word 的 ParagraphFormat.Alignment 取值 0=左、1=中、2=右。
    ; 方案里存的是“居左/居中/居右”文案，这里做一次显式映射，
    ; 不要依赖下拉下标，改文案时不会串位。
    alignVal := 1
    if (HeaderAlign == "居左")
        alignVal := 0
    else if (HeaderAlign == "居右")
        alignVal := 2

    ; 字体字号：字号走 ParseFontSize（“四号” → 14 磅），
    ; 加粗字段来自方案，统一转成布尔，避免注册表里读到字符串 "0"。
    hFont := HeaderFont
    if (hFont == "")
        hFont := ActiveStyle.BodyFont
    hSize := ParseFontSize(HeaderSize, 14.0)
    hBold := (Integer(HeaderBold) != 0)
    ; v6.8.24：不画文字下划线，横线只走段落下边框。v6.8.23 已改为：
    ; 不用 Font.Underline 的原因：那是字符级格式，线只跟着文字走，文字短就只
    ; 画一小截，居左/居右时也不会左右撑满；段落下边框是段落级格式，只要段落
    ; 左右缩进为 0（本函数后面会归零），线就自动铺满整个页眉版心宽度，
    ; 与文字长短、对齐方式无关。
    hUnderline := (Integer(HeaderUnderline) != 0)

    warnings := []

    ; 逐节处理：附件另起一页会产生新节，必须在分节完成之后再写页眉，
    ; 否则新节继承不到；本函数在 SetPageNumbers() 之前调用，此时节已定型。
    for section in doc.Sections {
        ; WPS 下 PageSetup 偶尔取不到，取不到就按“未启用首页/奇偶页”处理，
        ; 不影响主页眉的写入。
        ps := ""
        try ps := section.PageSetup

        ; 页眉类型：1=主页眉（每节必有）、2=首页页眉、3=偶数页页眉
        headerTypes := [1]
        if (IsObject(ps)) {
            try {
                if (ps.DifferentFirstPageHeaderFooter)
                    headerTypes.Push(2)
            }
            try {
                if (ps.OddAndEvenPagesHeaderFooter)
                    headerTypes.Push(3)
            }
        }

        for hType in headerTypes {
            try {
                header := section.Headers(hType)

                ; 先断开“链接到前一节”。注意顺序不能反：先写内容再断开，
                ; 写进去的内容会同时出现在上一节。
                try header.LinkToPrevious := false

                ; 记录原页眉是否有内容（文字/图片/表格），仅用于“添加·替换”时
                ; 在日志里提示“原有内容已被替换”，不改变处理方式。
                hadContent := false
                try {
                    oldText := header.Range.Text
                    oldText := StrReplace(oldText, Chr(13), "")
                    oldText := StrReplace(oldText, Chr(10), "")
                    oldText := StrReplace(oldText, Chr(12), "")
                    hadContent := (TrimAllSpaces(oldText) != ""
                        || header.Range.InlineShapes.Count > 0
                        || header.Range.Tables.Count > 0)
                }

                if (mode == 2) {
                    ; “清空”：先清文字，再逐个删域。WPS 下 Range.Text := ""
                    ; 有时清不干净，残留的 PAGE 域会让页眉继续显示旧页码。
                    try header.Range.Text := ""
                    try {
                        fields := header.Range.Fields
                        i := fields.Count
                        while (i >= 1) {
                            try fields.Item(i).Delete()
                            i -= 1
                        }
                    }
                    ; v6.8.23：“清空”时也要把上一版留下的段落下边框去掉，
                    ; 否则文字没了、线还在，页眉上会残留一条横线。
                    ApplyHeaderBottomBorder(header, false)
                } else {
                    ; “添加·替换”：整段替换成方案文字。直接给 Range.Text 赋值
                    ; 会把原有图片、表格、域一起清掉（用户已确认“直接替换”）。
                    ; 若原页眉含表格，个别 Word/WPS 版本会拒绝整体赋值，
                    ; 此时退回“先删除内容、再写入文字”。
                    try {
                        header.Range.Text := headerText
                    } catch {
                        header.Range.Delete()
                        header.Range.Text := headerText
                    }

                    ; 取页眉的第一个段落写格式：赋值后通常只剩一个段落。
                    hPara := header.Range.Paragraphs.Item(1)
                    try {
                        hPara.Range.Font.Name      := hFont
                        hPara.Range.Font.NameFarEast := hFont
                        hPara.Range.Font.NameAscii := hFont
                        hPara.Range.Font.Size      := hSize
                        hPara.Range.Font.Bold      := hBold
                        ; v6.8.23：显式清掉字符下划线。Range.Text 赋值会继承原页眉
                        ; 第一个字符的格式，原页眉若本来就带下划线，新文字会跟着带，
                        ; 再叠上段落下边框就成了“双线”。这里统一归零，
                        ; 横线只由下面的段落下边框负责。
                        hPara.Range.Font.Underline := 0
                    }
                    try hPara.Format.Alignment := alignVal
                    ; 页眉段落不能继承正文的样式：首行缩进、左右缩进、
                    ; 段前段后、行距全部归零，否则页眉会缩进两个字符或被撑高。
                    try hPara.Format.FirstLineIndent := 0
                    try hPara.Format.CharacterUnitFirstLineIndent := 0
                    try hPara.Format.LeftIndent := 0
                    try hPara.Format.CharacterUnitLeftIndent := 0
                    try hPara.Format.RightIndent := 0
                    try hPara.Format.CharacterUnitRightIndent := 0
                    try hPara.Format.SpaceBefore := 0
                    try hPara.Format.SpaceAfter := 0
                    try {
                        ; 0 = 单倍行距（页眉只有一行，不需要固定值 28 磅）
                        hPara.Format.LineSpacingRule := 0
                        hPara.Format.LineSpacing := 0
                    }

                    ; v6.8.23：加/去段落下边框。必须放在上面“左右缩进归零”之后：
                    ; 缩进不为 0 时下边框只画在缩进之间，撑不满整行。
                    ApplyHeaderBottomBorder(header, hUnderline)
                }

                if (mode == 1 && hadContent)
                    warnings.Push("第 " section.Index " 节" HeaderTypeLabel(hType) "原有内容已被替换")
            } catch Error as err {
                warnings.Push("第 " section.Index " 节" HeaderTypeLabel(hType) "写入失败：" err.Message)
            }
        }
    }

    if (warnings.Length == 0)
        return ""
    warnText := ""
    for w in warnings {
        if (warnText != "")
            warnText .= "；"
        warnText .= w
    }
    return warnText
}

; 给页眉段落加/去“下边框”，用来画一条横跨整个段落的页眉横线。v6.8.24 改名。
; 参数 enabled：true = 单实线，false = 去掉边框。
;
; COM 细节：
;   - Borders.Item(-3) 里的 -3 是 wdBorderBottom（WdBorderType 枚举：上 -1、
;     左 -2、下 -3、右 -4），别写成 3。
;   - LineStyle：0 = wdLineStyleNone（无线）、1 = wdLineStyleSingle（单实线）。
;   - LineWidth 只认 wdLineWidth 枚举值（2/4/6/8/12/18/24/36），
;     6 = 0.75 磅，接近国标页眉分隔线的粗细；不要直接写 0.75。
;   - Color：0 = 黑色（wdColorBlack）。
;
; 为什么用 header.Range.Borders 而不是 hPara.Format.Borders：
;   Range.Borders 会作用到范围内的所有段落，页眉里万一有多个段落也能一次处理；
;   个别 WPS 版本 Range.Borders 取不到时，再退回逐段设置 ParagraphFormat.Borders。
ApplyHeaderBottomBorder(header, enabled) {
    lineStyle := enabled ? 1 : 0
    done := false
    try {
        bd := header.Range.Borders.Item(-3)
        bd.LineStyle := lineStyle
        if (enabled) {
            bd.LineWidth := 6
            bd.Color := 0
        }
        done := true
    }
    if (done)
        return
    ; 兜底：逐段设置（Paragraphs 是 1 基集合，for 直接遍历）
    try {
        for para in header.Range.Paragraphs {
            bd2 := para.Format.Borders.Item(-3)
            bd2.LineStyle := lineStyle
            if (enabled) {
                bd2.LineWidth := 6
                bd2.Color := 0
            }
        }
    }
}

; 页眉类型的中文名（用于日志提示，与 FooterTypeLabel 对称）
HeaderTypeLabel(type) {
    if (type == 2)
        return "首页页眉"
    if (type == 3)
        return "偶数页页眉"
    return "页眉"
}

SetPageNumbers(doc, formatStr := "格式数字 — 1 — 居中", skipSinglePage := 1) {
    global ActiveStyle, gAttachSectionIndex
    warnings := []
    ; v6.8.14：页码字体字号统一取自方案（默认 宋体 四号 不加粗）
    pageFont := ActiveStyle.PageNumFont
    pageSize := ParseFontSize(ActiveStyle.PageNumSize, 14.0)
    pageBold := ActiveStyle.PageNumBold
    totalPages := doc.ComputeStatistics(2)
    disable := ((totalPages <= 1 && skipSinglePage) || InStr(formatStr, "无"))

    for section in doc.Sections {
        isAttachSec := (gAttachSectionIndex != 0 && section.Index == gAttachSectionIndex)
        pageSetup := section.PageSetup

        footerTypes := [1]
        if (pageSetup.DifferentFirstPageHeaderFooter)
            footerTypes.Push(2)
        if (pageSetup.OddAndEvenPagesHeaderFooter)
            footerTypes.Push(3)

        for footerType in footerTypes {
            bmName := "DocFmt_PN_S" section.Index "_T" footerType "_" A_TickCount
            try {
                footer := section.Footers(footerType)
                ; v6.8.15：附件节先清空页脚（避免顺延正文页码），再断开与上一节的链接，
                ; 避免清空操作把 LinkToPrevious 重置回 true。
                ; WPS 下 Range.Text := "" 有时清不干净，残留旧 PAGE 域会让页码继续顺延，
                ; 所以再逐个删除页脚里的所有域，确保后面写入的是全新的页码域。
                if (isAttachSec) {
                    try footer.Range.Text := ""
                    try {
                        fields := footer.Range.Fields
                        i := fields.Count
                        while (i >= 1) {
                            try fields.Item(i).Delete()
                            i -= 1
                        }
                    }
                }
                footer.LinkToPrevious := false

                ; v6.8.15（WPS 兼容）：断开链接后再设置页码从 1 开始。
                ; WPS 的 PageSetup 对象没有 PageNumbering 属性，必须用
                ; Footers(x).PageNumbers；Word 两者都支持，再补一次 PageSetup 兜底。
                if (isAttachSec) {
                    try {
                        pnums := footer.PageNumbers
                        pnums.RestartNumberingAtSection := true
                        pnums.StartingNumber := 1
                    }
                    try {
                        pn := pageSetup.PageNumbering
                        pn.RestartNumberingAtSection := true
                        pn.StartingNumber := 1
                    }
                }

                ; 先移除上一次写入的页码块，保留页脚其他内容
                removed := RemoveFooterPageBlock(footer)
                UpdateFooterPageFields(footer)

                if (disable) {
                    if (removed)
                        continue
                    ; 仅在页脚确实“只含页码”时清除，避免误删其他页脚内容
                    if (IsFooterOnlyPageNumber(footer))
                        footer.Range.Text := ""
                    else if (FooterHasPageField(footer))
                        warnings.Push("第 " section.Index " 节" FooterTypeLabel(footerType) "含其他内容，未自动移除页码，请人工检查")
                    continue
                }

                ; 页脚只含旧版页码：可安全整体替换
                if (IsFooterOnlyPageNumber(footer)) {
                    footer.Range.Text := ""
                    WriteFooterPageBlock(footer, 1, formatStr, bmName)
                    continue
                }

                ; 页脚已有页码且还有其他内容：保留原内容，只统一页码域字体
                if (FooterHasPageField(footer)) {
                    fields := footer.Range.Fields
                    Loop fields.Count {
                        f := fields.Item(A_Index)
                        if (f.Type == 33 || f.Type == 26)
                            SetRangeFont(f.Result, pageFont, pageSize, pageBold)
                    }
                    warnings.Push("第 " section.Index " 节" FooterTypeLabel(footerType) "含其他内容，已保留原有页码和页脚内容")
                    continue
                }

                visibleText := GetFooterVisibleText(footer)
                if (visibleText == "" && footer.Range.Tables.Count == 0) {
                    ; 页脚为空（或仅含图片等无文字内容），在第一个段落写入
                    WriteFooterPageBlock(footer, 1, formatStr, bmName)
                } else {
                    ; 页脚有其他内容但没有页码：在末尾追加独立段落，保留原内容
                    lastIdx := footer.Range.Paragraphs.Count
                    lastPara := footer.Range.Paragraphs.Item(lastIdx)
                    lastText := GetFooterVisibleText(lastPara)
                    lastParaHasObjects := (lastPara.Range.Bookmarks.Count > 0 || lastPara.Range.InlineShapes.Count > 0 || footer.Range.Tables.Count > 0)
                    try {
                        if (lastPara.Range.ShapeRange.Count > 0 || lastPara.Range.ContentControls.Count > 0)
                            lastParaHasObjects := true
                    }
                    if (lastText != "" || lastParaHasObjects) {
                        ; 在最后一个段落标记之前插入新段落，确保页码位于原有内容之后。
                        ; 不能直接对 footer.Range 调用 InsertParagraphAfter，否则单段页脚
                        ; 不会新增段落，页码会写到原段落开头。
                        endRange := footer.Range.Duplicate
                        endRange.SetRange(endRange.End - 1, endRange.End - 1)
                        endRange.InsertParagraphAfter()
                        lastIdx := footer.Range.Paragraphs.Count
                    }
                    WriteFooterPageBlock(footer, lastIdx, formatStr, bmName)
                }
            } catch Error as err {
                warnings.Push("第 " section.Index " 节" FooterTypeLabel(footerType) "处理失败：" err.Message)
            }
        }
    }

    if (warnings.Length == 0)
        return ""
    warnText := ""
    for w in warnings {
        if (warnText != "")
            warnText .= "；"
        warnText .= w
    }
    return warnText
}

ExportToPdf(doc, doc_path) {
    SplitPath(doc_path, , &dir, , &name_no_ext)
    pdf_path := dir "\" name_no_ext ".pdf"
    try {
        if (InStr(FileExist(pdf_path), "D"))
            return {ok: false, path: pdf_path, err: "同名 PDF 路径被文件夹占用"}
        if FileExist(pdf_path) {
            ; 先检测同名 PDF 是否被占用，避免导出时静默失败
            try {
                f := FileOpen(pdf_path, "a")
                f.Close()
            } catch {
                return {ok: false, path: pdf_path, err: "同名 PDF 已存在且被占用，请关闭后重试"}
            }
        }
        doc.ExportAsFixedFormat(pdf_path, 17, false, 0, 0)
        if (!FileExist(pdf_path))
            return {ok: false, path: pdf_path, err: "导出命令已执行，但未生成 PDF 文件"}
        if (FileGetSize(pdf_path) <= 0)
            return {ok: false, path: pdf_path, err: "生成的 PDF 文件为空"}
        return {ok: true, path: pdf_path, err: ""}
    } catch Error as err {
        return {ok: false, path: pdf_path, err: err.Message}
    }
}

FormatInlineHeading(doc, p, raw_text, hFont, hSize, hBold, bFont, bSize, bBold, maxLen, keepNext := false, standaloneAlign := 3, hStyle := 0) {
    ; v6.8.28：hStyle 为 0 表示不应用段落样式；非 0 时仅对“整段都是标题”的独立标题段落应用
    ; （如 -2=标题1）。标题与正文混排在同一段时不改段落样式，避免把正文也纳入标题样式/大纲。
    total_len := StrLen(raw_text)

    if (total_len <= maxLen) {
        if (hStyle != 0)
            TrySetParaStyle(p, hStyle)
        SetRangeFont(p.Range, hFont, hSize, hBold)
        p.Format.Alignment := standaloneAlign
        p.Format.CharacterUnitFirstLineIndent := 2
        if (keepNext)
            p.Format.KeepWithNext := true
        return
    }

    split_pos := RegExMatch(raw_text, "[。！？]")
    if (split_pos > 0 && split_pos <= maxLen) {
        if (hStyle != 0)
            TrySetParaStyle(p, -1)  ; 混排段落保持正文样式，只靠直接格式区分标题/正文片段
        SetRangeFont(doc.Range(p.Range.Start, p.Range.Start + split_pos), hFont, hSize, hBold)
        if (split_pos < total_len)
            SetRangeFont(doc.Range(p.Range.Start + split_pos, p.Range.End), bFont, bSize, bBold)
        p.Format.Alignment := 3
        p.Format.CharacterUnitFirstLineIndent := 2
        if (keepNext)
            p.Format.KeepWithNext := true
    } else {
        if (hStyle != 0)
            TrySetParaStyle(p, -1)
        SetRangeFont(p.Range, bFont, bSize, bBold)
        p.Format.Alignment := 3
        p.Format.CharacterUnitFirstLineIndent := 2
    }
}

; =====================================================================
; [文档网格] 每页行数 / 每行字符数（由方案字段驱动）
; =====================================================================
; 规则（v6.8.16 引入，v6.8.17 改为方案字段驱动）：
;   网格参数属于“整体配置方案”，和边距、页码、字体一起保存在 Profiles\<方案名> 下：
;     GridMode  0 = 不指定网格，1 = 指定行和字符网格，2 = 只指定行网格
;     GridLines 每页行数（GridMode 为 1 或 2 时生效）
;     GridChars 每行字符数（仅 GridMode = 1 时生效；0 = 不指定）
;   内置方案“党政机关公文格式”的默认值是 1 / 22 / 28，对应 GB/T 9704-2012 的版心；
;   其他方案的默认值是 2 / 44 / 0，与 Word「页面设置 → 文档网格」的默认口径一致。
;   这些默认值由 ProfileRead() 在旧方案（没有这三个键）里按边距预设推算得出。
;
; 实现要点（都是实测结论，改动前请先看，否则很容易踩坑）：
;   * 必须在设置页边距之后再调用。实测改完上下左右边距，Word/WPS 会按新的版心
;     重新推算网格（例如换成公文边距后自动变成 40 行 / 42 字符），先设网格会被它覆盖。
;   * doc.PageSetup 是整篇文档级别的属性，赋值会自动作用到所有分节，
;     不需要循环 doc.Sections 逐个设置，逐个设置反而更容易在 WPS 上出错。
;   * LayoutMode（WdLayoutMode 枚举）：0 = 无网格，1 = 指定行和字符网格，2 = 只指定行网格。
;   * “只指定行网格”绝对不能碰 CharsLine：实测只要给 CharsLine 赋任何值，Word 就会把
;     LayoutMode 自动掰回 1（变成“指定行和字符网格”）。文档里残留的旧 CharsLine
;     （例如上一次按公文版心排过的 28）不用清，LayoutMode = 2 下它被 Word 忽略。
;   * 需要写 CharsLine 时，赋值顺序固定为 LayoutMode → LinesPage → CharsLine。
;   * LinesPage 要显式赋值，不能指望 Word 按边距自动推算：实测换边距后它会自己算成
;     40 行之类，只有显式赋值才能保证界面里看到的“每页 N 行”。
;   * 绝对不能把 0 写进 LinesPage。实测会直接抛 RPC_E_SERVERFAULT (0x80010105)，
;     打断 Word/WPS 的 COM 连接，导致整篇文档排版失败；下面用 Max() 兜底。
;   * 每个属性单独 try：WPS 对个别网格属性可能不支持或报错，不能让一个网格设置
;     失败拖垮整篇排版（与本文件其它 COM 赋值的容错风格保持一致）。
;   * 只动页面级网格，不改段落行距。本程序的正文行距是“固定值 28 磅”，会覆盖网格
;     的行距效果；网格实际控制的是每页行数上限与每行字数上限。
ApplyDocumentGrid(doc) {
    global GridMode, GridLines, GridChars

    ; 字段可能来自手工改过的注册表，先做一次合法性兜底，避免把非法值写进 COM
    mode := Integer(GridMode)
    if (mode != 0 && mode != 2)
        mode := 1                       ; 非法值一律按“指定行和字符网格”处理
    lines := Max(1, Integer(GridLines))  ; 行数最少 1，绝不能是 0（会让 COM 崩溃）
    chars := Max(0, Integer(GridChars))  ; 0 表示不指定每行字符数

    try page_setup := doc.PageSetup
    catch
        return false        ; 文档已关闭或 COM 失效时直接跳过，不影响后续流程

    switch mode {
        case 0:
            ; 不指定网格：只关掉网格，行数/字数交给 Word 按段落行距自动排版
            try page_setup.LayoutMode := 0
        case 2:
            ; 只指定行网格：只写每页行数，**不写** CharsLine（见上方说明）
            try page_setup.LayoutMode := 2
            try page_setup.LinesPage  := lines
        default:
            ; 指定行和字符网格：按 LayoutMode → LinesPage → CharsLine 的顺序写
            try page_setup.LayoutMode := 1
            try page_setup.LinesPage  := lines
            if (chars > 0)
                try page_setup.CharsLine := chars
    }

    return true
}

; =====================================================================
; [核心排版引擎] 单文档格式化主流程
; =====================================================================
FormatWordDocument(file_path, app, &err_msg := "", &warn_msg := "") {
    global MarginTop, MarginBottom, MarginLeft, MarginRight, PageNumFormat, SkipSinglePageNum
    global ActiveStyle, LINE_SPACING_RULE, LINE_SPACING_BODY, LINE_SPACING_TITLE
    global ExportPDF, CreateBAK, OptPreventOrphan, OptAttachmentNewPage
    global BakDir, DefaultBakDir, MaxHeadingLen

    err_msg := "", warn_msg := ""

    if (!FileExist(file_path)) {
        err_msg := "文件不存在或无访问权限"
        return false
    }

    try {
        f := FileOpen(file_path, "rw")
        if (f)
            f.Close()
    } catch {
        err_msg := "文件被占用（可能正被 Word 或其他程序打开）"
        return false
    }

    if (CreateBAK) {
        targetDir := (BakDir != "") ? BakDir : DefaultBakDir
        if (!DirExist(targetDir)) {
            try DirCreate(targetDir)
        }
        SplitPath(file_path, , , &fileExt, &nameNoExt)
        bak_path := targetDir "\" nameNoExt "_" FormatTime(, "yyyyMMdd_HHmmss") "." fileExt
        try FileCopy(file_path, bak_path, true)
    }

    doc := ""
    try {
        try app.ScreenUpdating := false
        doc := app.Documents.Open(file_path, false, false, false)

        ptMTitle   := ParseFontSize(ActiveStyle.MTitleSize, 22.0)
        ptSubTitle := ParseFontSize(ActiveStyle.SubTitleSize, 16.0)
        ptAttachMark := ParseFontSize(ActiveStyle.AttachMarkSize, 16.0)  ; 附件标记“附件1”，国标 3 号
        ptAuthor   := ParseFontSize(ActiveStyle.AuthorSize, 16.0)
        ptBody     := ParseFontSize(ActiveStyle.BodySize, 16.0)
        ptH1       := ParseFontSize(ActiveStyle.H1Size, 16.0)
        ptH2       := ParseFontSize(ActiveStyle.H2Size, 16.0)
        ptH3       := ParseFontSize(ActiveStyle.H3Size, 16.0)
        ptH4       := ParseFontSize(ActiveStyle.H4Size, 16.0)
        ptTable    := ParseFontSize(ActiveStyle.TableSize, 10.5)

        try doc.Content.ListFormat.ConvertNumbersToText()

        page_setup := doc.PageSetup
        page_setup.TopMargin      := MmToPt(MarginTop)
        page_setup.BottomMargin   := MmToPt(MarginBottom)
        page_setup.LeftMargin     := MmToPt(MarginLeft)
        page_setup.RightMargin    := MmToPt(MarginRight)
        ; v6.8.21：只有“添加·替换”才写页眉距离；“不改变”“清空”完全不碰这一项，
        ; 否则会把用户原有页眉的位置也一起挪动（v6.8.20 及以前固定写 15 mm）。
        ; 注意：HeaderDistance 是“页眉顶端到页面上边”的距离，单位磅；
        ; 它只影响页眉位置，页眉段落自身的段前距/行距在 SetDocumentHeaders() 里已归零。
        if (HeaderMode == 1)
            page_setup.HeaderDistance := MmToPt(HeaderMarginTop)
        page_setup.FooterDistance := MmToPt(28)

        ; 文档网格必须在上面四项页边距赋值之后再设置：改动边距会让 Word/WPS
        ; 按新版心重新推算每页行数与每行字符数，先设网格会被边距覆盖。
        ; 具体规则与坑见 ApplyDocumentGrid 上方注释。
        ApplyDocumentGrid(doc)

        CleanAllBlankParagraphs(doc)

        ; 表格保护：勾选时先记录表格原有格式，排版完成后原样恢复
        tableSnapshots := ""
        if (OptProtectTable)
            tableSnapshots := CaptureTableFormats(doc)

        content := doc.Content
        ; v6.8.29：先把文档内置样式（正文/标题/副标题/标题1~4）的段前、段后、
        ; 行距调整为国标要求，避免后续应用样式时继承样式自带的 8 pt 段后距、
        ; 单倍行距等非国标间距。
        AdjustBuiltinStyleSpacing(doc)

        ; v6.8.28：先把全文基线设为“正文”(wdStyleNormal=-1)，建立段落样式结构；
        ; 随后立即用本方案的直接格式覆盖，避免样式默认值改变排版效果。
        try content.Style := -1
        SetRangeFont(content, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
        content.ParagraphFormat.Alignment := 3
        content.ParagraphFormat.CharacterUnitLeftIndent := 0
        content.ParagraphFormat.LeftIndent := 0
        content.ParagraphFormat.CharacterUnitRightIndent := 0
        content.ParagraphFormat.RightIndent := 0
        content.ParagraphFormat.CharacterUnitFirstLineIndent := 2
        content.ParagraphFormat.LineSpacingRule := LINE_SPACING_RULE
        content.ParagraphFormat.LineSpacing     := LINE_SPACING_BODY
        content.ParagraphFormat.SpaceBeforeAuto := 0
        content.ParagraphFormat.SpaceAfterAuto  := 0
        content.ParagraphFormat.LineUnitBefore  := 0
        content.ParagraphFormat.LineUnitAfter   := 0
        content.ParagraphFormat.SpaceBefore     := 0
        content.ParagraphFormat.SpaceAfter      := 0
        content.ParagraphFormat.KeepWithNext    := false
        
        Loop doc.Paragraphs.Count {
            if (Mod(A_Index, 20) == 0)
                Sleep(-1)
            p_tmp := doc.Paragraphs(A_Index)
            if (p_tmp.Range.Tables.Count == 0)
                GetCleanText(p_tmp)
        }

        if (doc.Paragraphs.Count == 0) {
            try {
                doc.Save()
                doc.Saved := true
            }
            return true
        }

        Loop doc.Paragraphs.Count {
            if (A_Index > doc.Paragraphs.Count)
                break
            p_curr := doc.Paragraphs(A_Index)
            if (p_curr.Range.Tables.Count > 0)
                continue
            t_curr := TrimAllSpaces(RTrim(p_curr.Range.Text, "`r`n"))

            if (RegExMatch(t_curr, "^附件\s*[:：]")) {
                att_idx := A_Index
                if (att_idx > 1) {
                    p_prev := doc.Paragraphs(att_idx - 1)
                    if (p_prev.Range.Tables.Count == 0) {
                        t_prev := TrimAllSpaces(RTrim(p_prev.Range.Text, "`r`n"))
                        if (t_prev != "") {
                            p_curr.Range.InsertParagraphBefore()
                            p_blank_att := doc.Paragraphs(att_idx)
                            p_blank_att.Format.CharacterUnitFirstLineIndent := 0
                            p_blank_att.Format.FirstLineIndent := 0
                            p_blank_att.Format.LineSpacingRule := LINE_SPACING_RULE
                            p_blank_att.Format.LineSpacing     := LINE_SPACING_BODY
                            att_idx++
                            p_curr := doc.Paragraphs(att_idx)
                        }
                    }
                }

                if (SubStr(t_curr, -1) == "。")
                    t_curr := SubStr(t_curr, 1, StrLen(t_curr) - 1)

                if RegExMatch(t_curr, "^附件\s*[:：]\s*(\d+)\s*[\.．、·](?!\d)\s*(.*)$", &m_att1) {
                    t_curr := "附件：" . m_att1[1] . "." . LTrimAllSpaces(m_att1[2])
                    SetParagraphText(p_curr, t_curr)
                    wrapChars := 5.0 + GetVisualLength(m_att1[1] . ".")
                } else if RegExMatch(t_curr, "^附件\s*[:：]\s*(.*)$", &m_att0) {
                    t_curr := "附件：" . LTrimAllSpaces(m_att0[1])
                    SetParagraphText(p_curr, t_curr)
                    wrapChars := 5.0
                } else {
                    wrapChars := 5.0
                }

                SetRangeFont(p_curr.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                p_curr.Format.CharacterUnitLeftIndent := 0
                p_curr.Format.CharacterUnitFirstLineIndent := 0
                p_curr.Format.LeftIndent := wrapChars * ptBody
                p_curr.Format.FirstLineIndent := (2.0 - wrapChars) * ptBody
                p_curr.Format.Alignment := 0

                next_idx := att_idx + 1
                while (next_idx <= doc.Paragraphs.Count) {
                    p_sub := doc.Paragraphs(next_idx)
                    if (p_sub.Range.Tables.Count > 0)
                        break
                    t_sub := TrimAllSpaces(RTrim(p_sub.Range.Text, "`r`n"))
                    if RegExMatch(t_sub, "^(\d+)\s*[\.．、·](?!\d)\s*(.*)$", &m_sub) {
                        sub_title := LTrimAllSpaces(m_sub[2])
                        if (SubStr(sub_title, -1) == "。")
                            sub_title := SubStr(sub_title, 1, StrLen(sub_title) - 1)

                        num_prefix := m_sub[1] . "."
                        t_sub := num_prefix . sub_title
                        SetParagraphText(p_sub, t_sub)

                        prefixLen := GetVisualLength(num_prefix)
                        SetRangeFont(p_sub.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                        p_sub.Format.CharacterUnitLeftIndent := 0
                        p_sub.Format.CharacterUnitFirstLineIndent := 0
                        p_sub.Format.LeftIndent := (5.0 + prefixLen) * ptBody
                        p_sub.Format.FirstLineIndent := -(prefixLen * ptBody)
                        p_sub.Format.Alignment := 0
                        next_idx++
                    } else {
                        break
                    }
                }
                break
            }
        }

        ; -----------------------------------------------------------------
        ; 标题区：先确定“大标题”，再从它下面连续收集副标题（v6.8.12）
        ;   · 默认第一个非空段就是大标题；
        ;   · 若第一段只是“附件1”“附件 1”这类单独成行的附件编号，
        ;     说明它是附件封面标记，大标题在紧随其后的段落；v6.8.13 起该
        ;     标记按“黑体三号、顶格”排版，并与标题之间补一个空行；
        ;   · 从大标题下一段开始向下扫描，末尾没有标点且不是一至四级标题
        ;     的段落全部作为副标题；遇到表格、标点结尾、各级标题、附件说明
        ;     或明显过长的正文即停止；
        ;   · 扫描时遇到的真空段直接删除，保证副标题之间没有空行；
        ;     含分页符/分节符的段落不算空段，保留并结束扫描（v6.8.13）。
        ; -----------------------------------------------------------------
        header_count := 0
        title_idx := 0
        ; 第一段是否为“附件1”这类附件标记（用于后面在标记与标题之间补空行）
        first_is_attachment_mark := false
        probe_limit := Min(doc.Paragraphs.Count, 5)
        Loop probe_limit {
            p_probe := doc.Paragraphs(A_Index)
            if (p_probe.Range.Tables.Count > 0)
                continue
            probe_text := TrimAllSpaces(StrReplace(RTrim(p_probe.Range.Text, "`r`n"), Chr(12), ""))
            if (probe_text == "")
                continue

            ; 单独成行的“附件1”“附件 1”“附件1：”是附件封面标记，不是标题。
            ; v6.8.13 起按 GB/T 9704-2012 7.3.8 排版：字体字号取方案的“附件标记”
            ; （默认黑体三号）、顶格、取消首行缩进；勾选“附件另起一页”时另面编排。
            if (RegExMatch(probe_text, "^附件\s*[0-9０-９一二三四五六七八九十]+\s*[:：]?\s*$")) {
                if (OptAttachmentNewPage)
                    ApplyAttachmentPageBreak(doc, p_probe, A_Index)
                TrySetParaStyle(p_probe, -1)
                SetRangeFont(p_probe.Range, ActiveStyle.AttachMarkFont, ptAttachMark, ActiveStyle.AttachMarkBold)
                p_probe.Format.Alignment := 0
                p_probe.Format.CharacterUnitFirstLineIndent := 0
                p_probe.Format.FirstLineIndent := 0
                first_is_attachment_mark := true
                continue
            }

            ; 以成文日期开头的文档没有独立大标题，保持原样，避免把日期放大居中
            if (IsDateText(probe_text))
                break

            title_idx := A_Index
            break
        }

        if (title_idx > 0) {
            ; 第一段是“附件1”标记时，先在标记和标题之间补一个空段：
            ; GB/T 9704-2012 7.3.8 要求附件标记占版心第一行、第二行留空、
            ; 附件标题居中于版心第三行。CleanAllBlankParagraphs 已清掉原有空段，
            ; 所以这里必须自己补一行，标题才会落在第三行。
            if (first_is_attachment_mark) {
                doc.Paragraphs(title_idx).Range.InsertParagraphBefore()
                p_gap_head := doc.Paragraphs(title_idx)
                p_gap_head.Format.CharacterUnitFirstLineIndent := 0
                p_gap_head.Format.FirstLineIndent := 0
                p_gap_head.Format.LineSpacingRule := LINE_SPACING_RULE
                p_gap_head.Format.LineSpacing     := LINE_SPACING_BODY
                title_idx += 1   ; 标题被挤到下一段
            }

            p_title := doc.Paragraphs(title_idx)
            title_text := TrimAllSpaces(StrReplace(RTrim(p_title.Range.Text, "`r`n"), Chr(12), ""))
            ; 大标题本身必须以非标点结尾，否则视为普通正文，不做标题区处理
            if (title_text != "" && !IsEndWithPunct(title_text)) {
                header_count := title_idx
                ; v6.8.28：主标题应用内置“标题”(wdStyleTitle=-63) 样式，便于大纲/导航识别；
                ; 设置样式后必须重新应用字体/段落格式，避免被样式默认值覆盖。
                TrySetParaStyle(p_title, -63)
                SetRangeFont(p_title.Range, ActiveStyle.MTitleFont, ptMTitle, ActiveStyle.MTitleBold)
                p_title.Format.Alignment := 1
                p_title.Format.CharacterUnitFirstLineIndent := 0
                p_title.Format.FirstLineIndent := 0
                p_title.Format.LineSpacing := LINE_SPACING_TITLE

                ; 从大标题下一段开始连续向下收副标题。
                ; sub_count 限制最多 10 段，防止极端文档整篇被误判成副标题。
                sub_idx := title_idx + 1
                sub_count := 0
                while (sub_idx <= doc.Paragraphs.Count && sub_count < 10) {
                    p_sub := doc.Paragraphs(sub_idx)
                    if (p_sub.Range.Tables.Count > 0)
                        break

                    sub_text := TrimAllSpaces(StrReplace(RTrim(p_sub.Range.Text, "`r`n"), Chr(12), ""))
                    if (sub_text == "") {
                        ; 含图片/图形的段落不是空行，不能删除，直接结束标题区
                        if (p_sub.Range.InlineShapes.Count > 0)
                            break
                        ; 含分页符/分节符或设置了段前分页的段落也不是空行，
                        ; 保留并结束标题区（v6.8.13，避免把分页符当空行删掉）
                        if (IsPageBreakProtected(p_sub))
                            break
                        ; 真空段直接删除；删除后下一段顶到当前索引，
                        ; 所以 sub_idx 不递增，继续检查同一位置。
                        try {
                            p_sub.Range.Delete()
                            continue
                        } catch {
                            sub_idx++
                            continue
                        }
                    }

                    ; 以标点结尾：主送机关、正文、附件说明等，标题区到此结束
                    if (IsEndWithPunct(sub_text))
                        break

                    ; 日期优先按副标题处理；其余段落若是一至四级标题则结束扫描
                    if (!IsDateText(sub_text) && IsHeadingText(sub_text))
                        break

                    ; “附件：1.xxx”是正文附件说明，不是副标题
                    if (RegExMatch(sub_text, "^附件\s*[:：]"))
                        break

                    ; 安全阀：副标题通常很短；超过 80 字的段落按正文处理，
                    ; 防止正文段末偶然没有标点时把整段改成居中的副标题。
                    if (StrLen(sub_text) > 80)
                        break

                    ; 统一使用当前方案的副标题字体（默认楷体三号、不加粗、居中）
                    ; v6.8.28：副标题尝试应用内置“副标题”(wdStyleSubtitle=-75)，WPS/旧模板不支持时静默降级
                    TrySetParaStyle(p_sub, -75)
                    SetRangeFont(p_sub.Range, ActiveStyle.SubTitleFont, ptSubTitle, ActiveStyle.SubTitleBold)
                    p_sub.Format.Alignment := 1
                    p_sub.Format.CharacterUnitFirstLineIndent := 0
                    p_sub.Format.FirstLineIndent := 0

                    header_count := sub_idx
                    sub_count++
                    sub_idx++
                }
            }
        }

        start_body_idx := 1
        if (header_count > 0) {
            r := doc.Paragraphs(header_count).Range
            r.Collapse(0)
            r.InsertParagraphAfter()

            blank_p := doc.Paragraphs(header_count + 1)
            blank_p.Format.CharacterUnitFirstLineIndent := 0
            blank_p.Format.FirstLineIndent := 0
            blank_p.Format.LineSpacingRule := LINE_SPACING_RULE
            blank_p.Format.LineSpacing     := LINE_SPACING_BODY
            start_body_idx := header_count + 2
        }

        if (start_body_idx <= doc.Paragraphs.Count) {
            body_range := doc.Range(doc.Paragraphs(start_body_idx).Range.Start, doc.Content.End)
            find_obj := body_range.Find
            find_obj.ClearFormatting()
            find_obj.Replacement.ClearFormatting()
            find_obj.Wrap := 0
            try find_obj.Execute("^l", , , , , , , 0, , "^p", 2)
        }

        next_is_attachment_title := false
        p_idx := start_body_idx

        while (p_idx <= doc.Paragraphs.Count) {
            if (Mod(p_idx, 15) == 0)
                Sleep(-1)

            p := doc.Paragraphs(p_idx)
            if (p.Range.Tables.Count > 0) {
                p_idx++
                continue
            }

            raw_text := RTrim(p.Range.Text, "`r`n")
            text := TrimAllSpaces(StrReplace(raw_text, Chr(12), ""))
            if (text == "") {
                p_idx++
                continue
            }

            if (RegExMatch(text, "^附件\s*\d*$")) {
                ; 附件标记“附件1”“附件2”（GB/T 9704-2012 7.3.8）：
                ; 3 号黑体、顶格，字体字号取方案里的“附件标记”；
                ; 勾选“附件另起一页”时让附件另面编排。
                if (OptAttachmentNewPage)
                    ApplyAttachmentPageBreak(doc, p, p_idx)

                TrySetParaStyle(p, -1)
                SetRangeFont(p.Range, ActiveStyle.AttachMarkFont, ptAttachMark, ActiveStyle.AttachMarkBold)
                p.Format.Alignment := 0
                p.Format.CharacterUnitFirstLineIndent := 0
                p.Format.FirstLineIndent := 0
                p.Format.LineSpacingRule := LINE_SPACING_RULE
                p.Format.LineSpacing     := LINE_SPACING_BODY

                ; 标记与附件标题之间补一个空段，让附件标题落在版心第三行。
                ; 例外：下一段只是分页符/分节符时不再补空行，否则空行会留在
                ; 上一页末尾，标题仍落在新页第一行，等于白占一行。
                p_next := (p_idx + 1 <= doc.Paragraphs.Count) ? doc.Paragraphs(p_idx + 1) : ""
                if (!IsObject(p_next) || !IsBreakOnlyParagraph(p_next)) {
                    r_mark := p.Range.Duplicate
                    r_mark.Collapse(0)
                    r_mark.InsertParagraphAfter()

                    p_gap := doc.Paragraphs(p_idx + 1)
                    p_gap.Format.CharacterUnitFirstLineIndent := 0
                    p_gap.Format.FirstLineIndent := 0
                    p_gap.Format.LineSpacingRule := LINE_SPACING_RULE
                    p_gap.Format.LineSpacing     := LINE_SPACING_BODY

                    next_is_attachment_title := true
                    p_idx += 2   ; 跳过标记段和新增空段，下一段就是附件标题
                } else {
                    next_is_attachment_title := true
                    p_idx++
                }
                continue
            }

            if (next_is_attachment_title) {
                ; v6.8.28：附件标题与主标题使用同级“标题”样式，保持文档大纲一致
                TrySetParaStyle(p, -63)
                SetRangeFont(p.Range, ActiveStyle.MTitleFont, ptMTitle, ActiveStyle.MTitleBold)
                p.Format.Alignment := 1
                p.Format.CharacterUnitFirstLineIndent := 0
                p.Format.FirstLineIndent := 0
                p.Format.LineSpacingRule := 4
                p.Format.LineSpacing     := LINE_SPACING_TITLE
                next_is_attachment_title := false

                r_att := p.Range
                r_att.Collapse(0)
                r_att.InsertParagraphAfter()

                p_blank := doc.Paragraphs(p_idx + 1)
                p_blank.Format.CharacterUnitFirstLineIndent := 0
                p_blank.Format.FirstLineIndent := 0
                p_blank.Format.LineSpacingRule := LINE_SPACING_RULE
                p_blank.Format.LineSpacing     := LINE_SPACING_BODY

                p_idx += 2
                continue
            }

            if (SubStr(text, 1, 3) == "【引】") {
                clean_quote := LTrimAllSpaces(SubStr(text, 4))
                SetParagraphText(p, clean_quote)
                SetRangeFont(p.Range, ActiveStyle.H2Font, ptH2, ActiveStyle.H2Bold)
                p.Format.Alignment := 3
                p.Format.CharacterUnitFirstLineIndent := 2
                p.Format.LineSpacingRule := LINE_SPACING_RULE
                p.Format.LineSpacing     := LINE_SPACING_BODY
                p_idx++
                continue
            }

            if (RegExMatch(text, "^附件\s*[:：]")) {
                p_idx++
                while (p_idx <= doc.Paragraphs.Count) {
                    p_sub := doc.Paragraphs(p_idx)
                    if (p_sub.Range.Tables.Count > 0)
                        break
                    t_sub := TrimAllSpaces(RTrim(p_sub.Range.Text, "`r`n"))
                    if (RegExMatch(t_sub, "^\d+\s*[\.．、·](?!\d)"))
                        p_idx++
                    else
                        break
                }
                continue
            }

            if (p_idx == start_body_idx && IsEndWithColon(text)) {
                ; v6.8.30：主送机关行套用自定义顶格段落样式“主送机关”（Alignment:=0、无首行缩进、无段前段后）。
                ; 样式不存在时自动创建；创建失败则降级为直接格式设置。设置样式后仍需重新施加
                ; 方案的正文字体/字号/加粗等直接格式，避免被样式默认字体覆盖。
                maStyle := GetOrCreateMainAddresseeStyle(doc)
                if (maStyle != "") {
                    try {
                        p.Range.Style := maStyle
                    } catch {
                        try {
                            p.Style := maStyle
                        }
                    }
                } else {
                    ; 降级：直接用正文样式并显式顶格
                    TrySetParaStyle(p, -1)
                }
                p.Format.Alignment := 0
                p.Format.CharacterUnitFirstLineIndent := 0
                p.Format.FirstLineIndent := 0
                SetRangeFont(p.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                p.Format.LineSpacingRule := LINE_SPACING_RULE
                p.Format.LineSpacing     := LINE_SPACING_BODY
                p.Format.SpaceBefore     := 0
                p.Format.SpaceAfter      := 0
                p_idx++
                continue
            }

            if (RegExMatch(text, "^([一二三四五六七八九十百]+)\s*([、.．·])\s*(.*)$", &m_h1) && !RegExMatch(m_h1[3], "^[一二三四五六七八九十百]+[、年]")) {
                clean_h1 := m_h1[1] . "、" . LTrimAllSpaces(m_h1[3])
                if (raw_text != clean_h1) {
                    SetParagraphText(p, clean_h1)
                    raw_text := clean_h1
                    text := clean_h1
                }
                ; v6.8.28：独立一级标题应用“标题 1”样式（-2），混排时由 FormatInlineHeading 保持正文样式
                FormatInlineHeading(doc, p, raw_text, ActiveStyle.H1Font, ptH1, ActiveStyle.H1Bold, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold, MaxHeadingLen, OptPreventOrphan, 0, -2)
            } else if RegExMatch(text, "^(\x28|（)\s*([一二三四五六七八九十百]+)\s*(\x29|）)\s*(.*)$", &m_h2) {
                clean_h2 := m_h2[1] . m_h2[2] . m_h2[3] . LTrimAllSpaces(m_h2[4])
                if (raw_text != clean_h2) {
                    SetParagraphText(p, clean_h2)
                    raw_text := clean_h2
                    text := clean_h2
                }
                FormatInlineHeading(doc, p, raw_text, ActiveStyle.H2Font, ptH2, ActiveStyle.H2Bold, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold, MaxHeadingLen, OptPreventOrphan, 3, -3)
            } else if RegExMatch(text, "^(\d+)\s*([.．、·])(?!\d)\s*(.*)$", &m_h3) {
                clean_h3 := m_h3[1] . "." . LTrimAllSpaces(m_h3[3])
                if (raw_text != clean_h3) {
                    SetParagraphText(p, clean_h3)
                    raw_text := clean_h3
                    text := clean_h3
                }
                FormatInlineHeading(doc, p, raw_text, ActiveStyle.H3Font, ptH3, ActiveStyle.H3Bold, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold, MaxHeadingLen, OptPreventOrphan, 3, -4)
            } else if RegExMatch(text, "^(\x28|（)\s*(\d+)\s*(\x29|）)\s*(.*)$", &m_h4) {
                clean_h4 := m_h4[1] . m_h4[2] . m_h4[3] . LTrimAllSpaces(m_h4[4])
                if (raw_text != clean_h4) {
                    SetParagraphText(p, clean_h4)
                    raw_text := clean_h4
                    text := clean_h4
                }
                FormatInlineHeading(doc, p, raw_text, ActiveStyle.H4Font, ptH4, ActiveStyle.H4Bold, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold, MaxHeadingLen, OptPreventOrphan, 3, -5)
            } else {
                ; v6.8.28：正文段落应用“正文”(Normal, -1) 样式，再覆盖为本方案的直接格式
                TrySetParaStyle(p, -1)
                SetRangeFont(p.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                p.Format.Alignment := 3
                p.Format.CharacterUnitFirstLineIndent := 2
                p.Format.LineSpacingRule := LINE_SPACING_RULE
                p.Format.LineSpacing     := LINE_SPACING_BODY
            }

            p_idx++
        }

        date_idx := 0, signer_idx := 0
        final_total := doc.Paragraphs.Count
        ; 只在标题区之后查找文末成文日期。
        ; v6.8.12 起，大标题下方的日期会被收进副标题块；如果仍从全文末尾
        ; 倒着找日期，这个副标题日期会被误当成落款日期重新右对齐。
        search_from := (start_body_idx > 0) ? start_body_idx : 1

        Loop final_total {
            idx := final_total - A_Index + 1
            if (idx < search_from)
                break
            p_chk := doc.Paragraphs(idx)
            if (p_chk.Range.Tables.Count > 0)
                continue
            txt := TrimAllSpaces(RTrim(p_chk.Range.Text, "`r`n"))
            if (RegExMatch(txt, "^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日") > 0) {
                date_idx := idx
                break
            }
        }

        if (date_idx > 0) {
            while (date_idx > 1) {
                p_above := doc.Paragraphs(date_idx - 1)
                if (p_above.Range.Tables.Count > 0)
                    break
                above_txt := TrimAllSpaces(RTrim(p_above.Range.Text, "`r`n"))
                if (RegExMatch(above_txt, "^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日") > 0) {
                    ; 这一段的末尾若带分页符/分节符（Chr(12)），删掉会丢页，直接放弃去重
                    if (IsPageBreakProtected(p_above))
                        break
                    try {
                        p_above.Range.Delete()
                        date_idx--
                    } catch {
                        break
                    }
                } else {
                    break
                }
            }

            cand_idx := date_idx - 1
            while (cand_idx >= search_from) {
                p_cand := doc.Paragraphs(cand_idx)
                if (p_cand.Range.Tables.Count > 0) {
                    cand_idx--
                    continue
                }
                cand_txt := TrimAllSpaces(RTrim(p_cand.Range.Text, "`r`n"))
                if (cand_txt == "") {
                    cand_idx--
                    continue
                }
                isDateLine := (RegExMatch(cand_txt, "^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日") > 0)
                if (!IsEndWithPunct(cand_txt) && !RegExMatch(cand_txt, "^(\d+[\.．、]|附件)") && !isDateLine && StrLen(cand_txt) <= 25)
                    signer_idx := cand_idx
                break
            }

            sign_top_idx := (signer_idx > 0) ? signer_idx : date_idx
            base_pages := doc.ComputeStatistics(2)

            doc.Paragraphs(sign_top_idx).Range.InsertParagraphBefore()
            p_blank1 := doc.Paragraphs(sign_top_idx)
            p_blank1.Format.CharacterUnitFirstLineIndent := 0
            p_blank1.Format.FirstLineIndent := 0
            p_blank1.Format.LineSpacingRule := LINE_SPACING_RULE
            p_blank1.Format.LineSpacing     := LINE_SPACING_BODY

            doc.Paragraphs(sign_top_idx + 1).Range.InsertParagraphBefore()
            p_blank2 := doc.Paragraphs(sign_top_idx + 1)
            p_blank2.Format.CharacterUnitFirstLineIndent := 0
            p_blank2.Format.FirstLineIndent := 0
            p_blank2.Format.LineSpacingRule := LINE_SPACING_RULE
            p_blank2.Format.LineSpacing     := LINE_SPACING_BODY

            if (doc.ComputeStatistics(2) > base_pages) {
                doc.Paragraphs(sign_top_idx + 1).Range.Delete()
                inserted_count := 1
            } else {
                inserted_count := 2
            }

            if (signer_idx > 0)
                signer_idx += inserted_count
            date_idx += inserted_count

            p_date := doc.Paragraphs(date_idx)
            clean_date := RegExReplace(TrimAllSpaces(RTrim(p_date.Range.Text, "`r`n")), "\s+", " ")

            clean_signer := ""
            if (signer_idx > 0) {
                p_signer := doc.Paragraphs(signer_idx)
                clean_signer := TrimAllSpaces(RTrim(p_signer.Range.Text, "`r`n"))
            }

            if (signer_idx > 0 && clean_signer != "") {
                lenSigner := GetVisualLength(clean_signer)
                lenDate   := GetVisualLength(clean_date)

                SetParagraphText(p_signer, clean_signer)
                SetParagraphText(p_date, clean_date)

                SetRangeFont(p_signer.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                SetRangeFont(p_date.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)

                if (lenSigner > lenDate) {
                    signerRightChars := 2.0
                    dateRightChars   := Round(2.0 + (lenSigner - lenDate) / 2.0, 1)
                } else {
                    dateRightChars   := 4.0
                    signerRightChars := Round(4.0 + (lenDate - lenSigner) / 2.0, 1)
                }

                p_signer.Format.FirstLineIndent := 0
                p_signer.Format.CharacterUnitFirstLineIndent := 0
                p_signer.Format.LeftIndent := 0
                p_signer.Format.CharacterUnitLeftIndent := 0
                p_signer.Format.RightIndent := 0
                p_signer.Format.CharacterUnitRightIndent := signerRightChars
                p_signer.Format.Alignment := 2

                p_date.Format.FirstLineIndent := 0
                p_date.Format.CharacterUnitFirstLineIndent := 0
                p_date.Format.LeftIndent := 0
                p_date.Format.CharacterUnitLeftIndent := 0
                p_date.Format.RightIndent := 0
                p_date.Format.CharacterUnitRightIndent := dateRightChars
                p_date.Format.Alignment := 2
            } else {
                SetParagraphText(p_date, clean_date)
                SetRangeFont(p_date.Range, ActiveStyle.BodyFont, ptBody, ActiveStyle.BodyBold)
                p_date.Format.FirstLineIndent := 0
                p_date.Format.CharacterUnitFirstLineIndent := 0
                p_date.Format.LeftIndent := 0
                p_date.Format.CharacterUnitLeftIndent := 0
                p_date.Format.RightIndent := 0
                p_date.Format.CharacterUnitRightIndent := 4.0
                p_date.Format.Alignment := 2
            }
        }

        RemoveTrailingEmptyParagraphs(doc)

        ; -----------------------------------------------------------------
        ; 步骤 8：后置处理表格字体、字号、段落对齐与缩进
        ; 仅在未勾选“表格保护与字号适配”时执行；勾选时表格保持原格式
        ; -----------------------------------------------------------------
        if (!OptProtectTable && doc.Tables.Count > 0) {
            for tbl in doc.Tables {
                try {
                    tblRange := tbl.Range
                    SetRangeFont(tblRange, ActiveStyle.TableFont, ptTable, ActiveStyle.TableBold)
                    tblParaFormat := tblRange.ParagraphFormat
                    tblParaFormat.CharacterUnitFirstLineIndent := 0
                    tblParaFormat.FirstLineIndent := 0
                    tblParaFormat.CharacterUnitLeftIndent := 0
                    tblParaFormat.LeftIndent := 0
                    tblParaFormat.LineSpacingRule := 0
                    tblParaFormat.SpaceBefore     := 0
                    tblParaFormat.SpaceAfter      := 0
                    tbl.Rows.Alignment := 1 ; 1 = wdAlignRowCenter（表格整体居中）
                    for cell in tbl.Range.Cells {
                        cell.VerticalAlignment := 1 ; 1 = wdCellAlignVerticalCenter（单元格垂直居中）
                        for cellPara in cell.Range.Paragraphs {
                            cellPara.Format.Alignment := 0 ; 0 = wdAlignParagraphLeft（单元格文字左对齐）
                        }
                    }
                }
            }
        }

        ; 勾选“表格保护”时，把表格原有格式原样写回
        if (OptProtectTable)
            RestoreTableFormats(doc, tableSnapshots)

        ; v6.8.21：先写页眉，再写页码。两步的提示都要保留——
        ; v6.8.20 及以前这里是直接 warn_msg := pageNumWarn，会把之前的提示覆盖掉。
        headerWarn := SetDocumentHeaders(doc)
        if (headerWarn != "")
            warn_msg := (warn_msg == "") ? headerWarn : warn_msg "；" headerWarn

        pageNumWarn := SetPageNumbers(doc, PageNumFormat, SkipSinglePageNum)
        if (pageNumWarn != "")
            warn_msg := (warn_msg == "") ? pageNumWarn : warn_msg "；" pageNumWarn
        global gAttachSectionIndex
        gAttachSectionIndex := 0

        try {
            doc.Save()
            doc.Saved := true
        } catch {
            err_msg := "保存失败，文件可能在排版中途被锁定或无权限"
            throw Error(err_msg)
        }

        if (ExportPDF) {
            pdfResult := ExportToPdf(doc, file_path)
            if (!pdfResult.ok) {
                pdfWarn := "PDF 导出失败（" pdfResult.path "）：" pdfResult.err
                warn_msg := (warn_msg == "") ? pdfWarn : warn_msg "；" pdfWarn
            }
        }

        try app.ScreenUpdating := true
        return true

    } catch Error as err {
        try app.ScreenUpdating := true
        err_msg := err.Message
        return false
    } finally {
        if (IsSet(doc) && doc != "") {
            try {
                doc.Saved := true
                doc.Close(0)
            }
        }
    }
}

; =====================================================================
; [帮助菜单：关于本程序 / 打赏作者弹窗]
; =====================================================================
ShowAboutDialog(*) {
    global mainGui, APP_VERSION
    abtGui := Gui("+Owner" mainGui.Hwnd " +ToolWindow", "请作者喝一杯咖啡")
    abtGui.SetFont("s14 bold c0055AA", "Microsoft YaHei")
    abtGui.Add("Text", "x20 y20 w410 Center", "公文自动排版引擎")
    
    abtGui.SetFont("s10 norm c000000")
    infoText := "版本：V" APP_VERSION "`n(国标 GB/T 9704-2012)`n`n"
             . "作者：Eyes`n"
             . "ＱＱ：28525626`n`n"
             . "如果这个小工具帮到了您，`n欢迎请作者喝杯咖啡支持一下！"
    abtGui.Add("Text", "x30 y75 w220", infoText)
    
    qrBase64 := "
    (
    
    )"
    
    qrCodePath := Base64ToTempFile(qrBase64)
    if (qrCodePath != "" && FileExist(qrCodePath)) {
        abtGui.Add("Picture", "x260 y70 w160 h-1", qrCodePath)
    } else {
        abtGui.SetFont("s9 c777777")
        abtGui.Add("Text", "x260 y140 w160 Center", "(收款码加载失败)")
        abtGui.SetFont("s10 c000000")
    }
    
    btnOK := abtGui.Add("Button", "x175 y280 w100 h32 Default", "确 定")
    CleanTempFile(*) {
        abtGui.Destroy()
        if (qrCodePath != "" && FileExist(qrCodePath)) {
            try FileDelete(qrCodePath)
        }
    }
    btnOK.OnEvent("Click", CleanTempFile)
    abtGui.OnEvent("Close", CleanTempFile)
    abtGui.Show("w450 h330")
}

Base64ToTempFile(base64Str) {
    tempPath := A_Temp "\DocFormatTool_QRCode_" A_TickCount ".jpg"
    base64Str := RegExReplace(base64Str, "\s") 
    if (base64Str == "")
        return ""
    size := 0
    DllCall("crypt32\CryptStringToBinary", "Str", base64Str, "UInt", 0, "UInt", 1, "Ptr", 0, "Ptr", 0)
    if (size == 0)
        return ""
    buf := Buffer(size)
    DllCall("crypt32\CryptStringToBinary", "Str", base64Str, "UInt", 0, "UInt", 1, "Ptr", buf, "Ptr", 0, "Ptr", 0)
    try {
        f := FileOpen(tempPath, "w")
        f.RawWrite(buf, size)
        f.Close()
        return tempPath
    } catch {
        return ""
    }
}