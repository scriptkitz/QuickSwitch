#Requires AutoHotkey >=2.0
#Singleinstance Force
; ============================================================================
; QuickSwitch V2
; ----------------------------------------------------------------------------
; 给系统的"打开/保存"文件对话框加一条贴着下边缘外侧的工具条：
;     [ ComboBox 搜索框 ] [ ⚙ 菜单 ]
;
;   * 有效路径 + 回车        -> 让对话框跳到该目录；文件则跳到所在目录并填好文件名
;   * 像路径但不存在         -> 搜索框红字"无效路径" + 错误提示音
;   * 其它输入 + 回车        -> 走 Everything 搜索，结果在独立 ListView 里列出，双击跳转
;   * 打字时实时搜索         -> 防抖 LIVE_DELAY 毫秒
;   * 自动切换               -> 对话框打开时按其历史目录 / 相邻资源管理器目录自动跳转
;
; 依赖：Everything3.dll（与脚本同目录）且 Everything 客户端在运行 —— 仅"搜索"功能需要；
;       没有它时退化为纯路径跳转，并在搜索时提示"Everything 未运行"。
; 模块：Everything3.ahk（Everything3/SDK3 封装）、Sqlite.ahk
;
; 目录
;   1 指令与常量   2 全局状态   3 主循环      4 工具条      5 搜索
;   6 结果列表     7 路径与对话框            8 配置        9 首选项
; ============================================================================

#Include "Sqlite.ahk"
#Include "Everything3.ahk"

SetWinDelay -1
SetControlDelay -1
SetWorkingDir(A_ScriptDir)

If (VerCompare(A_OSVersion, "10.0.0") <= 0) {
	MsgBox(A_OSVersion " is not supported.")
	ExitApp
}

; ============================================================================
; 1 指令与常量
; ============================================================================
; —— Win32 常量（只保留实际用到的）——
SWP_NOSIZE         := 0x0001
SWP_NOZORDER       := 0x0004
SWP_NOACTIVATE     := 0x0010
SWP_SHOWWINDOW     := 0x0040
SWP_HIDEWINDOW     := 0x0080
SWP_NOSENDCHANGING := 0x0400

EVENT_OBJECT_LOCATIONCHANGE := 0x800B
OBJID_WINDOW                := 0

; —— 工具条几何：[ ComboBox ] [ ⚙ ] 同一父窗口，贴在对话框下边缘外侧 ——
BAR_H     := 30
BAR_BTN_W := 32
BAR_GAP   := 4

; —— 文案 / 行为 ——
COMBO_PATH_PREFIX := "📁 "   ; 下拉项前缀：标记"最近文件夹路径"项
INVALID_TEXT      := "无效路径"
LIVE_DELAY        := 150     ; 实时搜索防抖（毫秒）
MAX_DIALOG_PATHS  := 15      ; 每个对话框保留的最近路径条数
HISTORY_MAX_DEF   := 20      ; 输入历史默认条数上限

; ============================================================================
; 2 全局状态
; ============================================================================
G_IniConfig := IniConfig("config.ini")
G_Menu := Menu()

; —— 当前对话框会话 ——
G_DlgHwnd     := 0
G_LastSetDlg := 0            ; 本会话已自动设置过路径的对话框，避免重复
G_DialogKey   := ""          ; "进程名___标题"，作为配置里的对话框标识
G_FDType      := ""          ; GetFileDialogType 的结果
G_bAutoSwitch := True        ; 本次是否自动切换路径
G_Jumping     := False       ; 跳转防重入（双击/回车重复触发只执行一次）

; —— 工具条 ——
G_BarGui      := 0
G_BarHwnd     := 0
G_Combo       := 0           ; ComboBox 控件对象
G_ComboHwnd   := 0
G_BarBtnHwnd  := 0
G_ComboWinH   := 0           ; 预计算：ComboBox 窗口总高（含下拉区）
G_ComboY      := 3           ; 预计算：ComboBox 纵向偏移
G_BtnY        := 3           ; 预计算：⚙ 按钮纵向偏移
G_BarChildW   := -1          ; 子控件上次按多宽排过版（-1 = 需重排）
G_ErrText     := ""          ; 当前以红字显示的提示文本

; —— 搜索结果列表（独立 ListView 窗口）——
G_ListGui   := 0
G_List      := 0             ; ListView 控件对象
G_ListHwnd  := 0             ; 列表窗口句柄（0 = 未显示）
G_ListItems := []            ; 行号 → 完整路径
G_ListH     := 66            ; 列表窗口高度（显示时按行数计算）
G_LastQuery := ""            ; 上次已搜过的关键词，避免重复查询

; —— 首选项控件 ——
G_PrefGui     := 0
G_PrefBar     := 0
G_PrefAutoRun := 0
G_PrefAutoSw  := 0
G_PrefMax     := 0

; —— 窗口事件钩子 ——
; 不全局注册，而是在检测到对话框时按"该对话框所在线程"注册，
; 只接收该线程的事件，避免被系统里无关窗口的移动/缩放唤醒。
G_WinEventProc := CallbackCreate(WinEventProc, "Fast", 7)
G_WinEventHook := 0

; —— 持久化设置 ——
G_bEnableBar          := G_IniConfig.ReadSetting("enable_bar", 1) ? True : False
G_bDefaultAutoSwitch  := G_IniConfig.ReadSetting("default_autoswitch", 1) ? True : False

; —— 主显示器工作区（判断结果列表"下方放不下则翻到对话框上方"）——
MonitorGetWorkArea(MonitorGetPrimary(), &WK_L, &WK_T, &WK_R, &WK_B)

G_BarGui := CreateBarWnd()
G_BarHwnd := G_BarGui.Hwnd

; ============================================================================
; 3 主循环：等一个文件对话框，挂上工具条，等会话结束
; ============================================================================
Loop {
	WinID := WinWaitActive("ahk_class #32770")
	G_FDType := GetFileDialogType(WinID)
	if (!G_FDType)
		continue

	G_DlgHwnd := WinID
	G_DialogKey := WinGetProcessName(WinID) . "___" . WinGetTitle(WinID)
	ShowBar(WinID)

	; 同一个对话框在失去激活前只自动设置一次路径
	if (G_LastSetDlg = G_DlgHwnd)
		continue

	_bAuto := G_IniConfig.ReadDialogNode(G_DialogKey)
	if (_bAuto == "") {                        ; 首次见到该对话框 → 记下默认值
		_bAuto := G_bDefaultAutoSwitch ? 1 : 0
		G_IniConfig.WriteDialogNode(G_DialogKey, _bAuto)
	}
	G_bAutoSwitch := _bAuto ? True : False

	FolderPath := GetZWindowPath(WinID)
	if (G_bAutoSwitch && ValidFolder(FolderPath))
		FeedDialog(WinID, FolderPath)

	WaitDialogSessionEnd(WinID)
	G_IniConfig.WriteDialogPaths(G_DialogKey, FolderPath)
	HideBar()
	G_LastSetDlg := 0
}
ExitApp(0)

; ============================================================================
; 4 工具条：[ ComboBox ] [ ⚙ ]，跟随对话框显示 / 移动 / 隐藏
; ============================================================================

; 创建工具条（只创建一次，之后只做显示/隐藏/移动）
CreateBarWnd() {
	global G_Combo, G_ComboHwnd, G_BarBtnHwnd
	global G_ComboWinH, G_ComboY, G_BtnY, BAR_H

	_gui := Gui("+ToolWindow -Caption +AlwaysOnTop")
	_gui.MarginX := 4
	_gui.MarginY := 3
	_gui.BackColor := "F2F2F2"

	; 搜索框：h200 是下拉列表高度；后续用 Add()/Delete() 放"最近路径 + 输入历史"
	G_Combo := _gui.AddComboBox("w280 h200", [])
	G_ComboHwnd := G_Combo.Hwnd
	G_Combo.OnEvent("Change", OnSearchInput)

	; ⚙ 菜单按钮（位置/尺寸由 RepositionBar 统一布局）
	_btn := _gui.AddButton("x300 y3 w32 h24", "⚙")
	G_BarBtnHwnd := _btn.Hwnd
	_btn.OnEvent("Click", OnBtnSetting)

	; 隐藏默认按钮：ComboBox 没有 KeyDown 事件，用它捕获回车
	_hidden := _gui.AddButton("Default Hidden w1 h1", "")
	_hidden.OnEvent("Click", OnSearchSubmit)

	; 预计算布局常量：只测一次，之后每次重定位不再 GetWindowRect/SendMessage
	_r := Buffer(16)
	DllCall("GetWindowRect", "Ptr", G_ComboHwnd, "Ptr", _r)
	G_ComboWinH := NumGet(_r, 12, "Int") - NumGet(_r, 4, "Int")     ; 含下拉区的窗口总高
	_vh := DllCall("SendMessage", "Ptr", G_ComboHwnd, "UInt", 0x0154, "Ptr", -1, "Ptr", 0, "Ptr")  ; CB_GETITEMHEIGHT(-1)：选择区可见高度
	if (!(_vh >= 16 && _vh <= BAR_H - 4))
		_vh := 22
	G_ComboY := (BAR_H - _vh) // 2
	G_BtnY := (BAR_H - (BAR_H - 6)) // 2
	return _gui
}

; 显示工具条：重建下拉、把位置钩子限定到该对话框线程、摆好位置
ShowBar(_WinID) {
	global G_BarGui, G_BarHwnd, G_BarChildW, G_bEnableBar, G_LastQuery
	if (!G_bEnableBar)
		return
	G_BarHwnd := G_BarGui.Hwnd
	G_BarChildW := -1            ; 强制重排子控件
	G_LastQuery := ""            ; 新会话：清掉"已搜过"记性
	RefreshComboItems(True)
	HookDialogThread(_WinID)
	RepositionBar(_WinID)
}

HideBar() {
	global G_BarHwnd
	UnhookDialogThread()
	HideResultList()
	if (G_BarHwnd && DllCall("IsWindow", "Ptr", G_BarHwnd))
		WinHideById(G_BarHwnd)
	G_BarHwnd := 0
}

; 重建下拉项：先"最近文件夹路径"(带前缀)，再"输入历史"
RefreshComboItems(clearText := False) {
	global G_Combo, G_IniConfig, G_DialogKey
	if (!G_Combo)
		return
	items := []
	for p in G_IniConfig.ReadDialogPaths(G_DialogKey)
		items.Push(COMBO_PATH_PREFIX . p)
	for k in G_IniConfig.ReadSearchHistory()
		items.Push(k)
	G_Combo.Delete()
	if (items.Length)
		G_Combo.Add(items)
	if (clearText)
		G_Combo.Text := ""
}

; 只监听指定窗口所属线程的位置/尺寸事件（跨进程用 out-of-context 钩子）
HookDialogThread(winID) {
	global G_WinEventHook, G_WinEventProc
	if (G_WinEventHook)
		DllCall("UnhookWinEvent", "Ptr", G_WinEventHook)
	tid := DllCall("GetWindowThreadProcessId", "Ptr", winID, "Ptr", 0, "UInt")
	G_WinEventHook := DllCall("SetWinEventHook", "UInt", EVENT_OBJECT_LOCATIONCHANGE, "UInt", EVENT_OBJECT_LOCATIONCHANGE, "Ptr", 0, "Ptr", G_WinEventProc, "UInt", 0, "UInt", tid, "UInt", 0, "Ptr")
}

UnhookDialogThread() {
	global G_WinEventHook
	if (G_WinEventHook) {
		DllCall("UnhookWinEvent", "Ptr", G_WinEventHook)
		G_WinEventHook := 0
	}
}

; 把工具条贴到对话框下边缘外侧（屏幕坐标）并排好内部控件。
; ⚠ 会被 Fast 回调调用：只用 DllCall/Buffer/NumGet。
RepositionBar(_WinID) {
	global G_BarHwnd, G_ComboHwnd, G_BarBtnHwnd
	global G_ComboWinH, G_ComboY, G_BtnY, G_BarChildW
	static _r := Buffer(16)
	if (!G_BarHwnd)
		return
	if (!DllCall("GetWindowRect", "Ptr", _WinID, "Ptr", _r))   ; 对话框已销毁
		return
	_l := NumGet(_r, 0, "Int")
	_b := NumGet(_r, 12, "Int")
	_w := NumGet(_r, 8, "Int") - _l
	if (_w < 200)
		_w := 200
	; 用 SWP_NOZORDER 保持其 AlwaysOnTop 层位，避免被对话框宿主盖住
	DllCall("SetWindowPos", "Ptr", G_BarHwnd, "Ptr", 0, "Int", _l, "Int", _b, "Int", _w, "Int", BAR_H, "UInt", SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW | SWP_NOSENDCHANGING, "Int")
	; 子控件是子窗口、随父窗自动移动；只有宽度变化（缩放）时才需要重排
	if (_w != G_BarChildW) {
		G_BarChildW := _w
		DllCall("SetWindowPos", "Ptr", G_ComboHwnd, "Ptr", 0, "Int", 4, "Int", G_ComboY, "Int", _w - 8 - BAR_BTN_W - BAR_GAP, "Int", G_ComboWinH, "UInt", SWP_NOZORDER | SWP_NOACTIVATE, "Int")
		DllCall("SetWindowPos", "Ptr", G_BarBtnHwnd, "Ptr", 0, "Int", _w - 4 - BAR_BTN_W, "Int", G_BtnY, "Int", BAR_BTN_W, "Int", BAR_H - 6, "UInt", SWP_NOZORDER | SWP_NOACTIVATE, "Int")
	}
	RepositionResultList()      ; 结果列表若在显示，跟着一起走
}

; 等待"对话框会话"结束：对话框被销毁，且焦点不在本工具自己的任何窗口上。
; 用 WinWaitNotActive（AHK 内置、开销极低）代替轮询；只有焦点落在本工具窗口上时才短暂轮询。
WaitDialogSessionEnd(WinID) {
	global G_BarHwnd, G_ListHwnd, G_PrefGui
	Loop {
		WinWaitNotActive(WinID)
		if (!DllCall("IsWindow", "Ptr", WinID))
			return
		if (G_BarHwnd && WinActive("ahk_id " G_BarHwnd)) {
			Sleep 30
			continue
		}
		if (G_ListHwnd && WinActive("ahk_id " G_ListHwnd)) {
			Sleep 30
			continue
		}
		if (G_PrefGui) {
			try {
				if (WinActive("ahk_id " G_PrefGui.Hwnd)) {
					Sleep 30
					continue
				}
			}
		}
		return
	}
}

; SetWinEventHook 回调：只处理当前对话框自身(OBJID_WINDOW)的位置/尺寸变化
WinEventProc(hHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime) {
	global G_DlgHwnd, G_BarHwnd
	if (event == EVENT_OBJECT_LOCATIONCHANGE && idObject == OBJID_WINDOW && hwnd == G_DlgHwnd && G_BarHwnd)
		RepositionBar(hwnd)
}

; ⚙ 菜单：自动切换 / 首选项
OnBtnSetting(Obj, info) {
	global G_BarHwnd, G_bAutoSwitch, G_IniConfig, G_DialogKey
	MENU_AUTOSWITCH := "自动切换"
	MENU_PREFS := "首选项"

	OnMenuPick(_name, _pos, _menu) {
		if (_name = MENU_AUTOSWITCH) {
			_menu.ToggleCheck(_name)
			G_IniConfig.WriteDialogNode(G_DialogKey, !G_bAutoSwitch)
		} else 		if (_name = MENU_PREFS) {
			SetTimer(ShowPreferences, -10)      ; 推迟到菜单关闭后再弹窗
		}
	}

	G_Menu.Delete()
	G_Menu.Add(MENU_AUTOSWITCH, OnMenuPick)
	G_bAutoSwitch ? G_Menu.Check(MENU_AUTOSWITCH) : G_Menu.UnCheck(MENU_AUTOSWITCH)
	G_Menu.Add()                                 ; 分隔线
	G_Menu.Add(MENU_PREFS, OnMenuPick)

	; 菜单挂到工具条窗口下，显示在 ⚙ 按钮正下方
	if (DllCall("GetParent", "Ptr", G_Menu.Handle) != G_BarHwnd)
		DllCall("SetParent", "Ptr", G_Menu.Handle, "Ptr", G_BarHwnd)
	ControlGetPos &cx, &cy, &cw, &ch, Obj
	G_Menu.Show(cx, cy + ch)
}

; ============================================================================
; 5 搜索：输入 / 回车 / 实时搜索 / 错误提示
; ============================================================================

; ComboBox Change：打字或从下拉选择都会触发
OnSearchInput(GuiCtrlObj, Info) {
	global LIVE_DELAY
	t := GuiCtrlObj.Text
	; 选中"最近文件夹路径"项 → 直接跳转并清空搜索框
	if (t != "" && IsPathItemText(t)) {
		if (JumpToPath(StripItemPrefix(t))) {
			GuiCtrlObj.Text := ""
			return
		}
	}
	; 防抖实时搜索：每次输入都重置一次性定时器，停手 LIVE_DELAY 毫秒才真正查一次
	SetTimer(LiveSearch, -LIVE_DELAY)
}

; 隐藏默认按钮的 Click = 搜索框里按了回车：
;   有效路径 → 跳转；像路径但不存在 → 红字报错；其余 → Everything 搜索并聚焦结果列表
OnSearchSubmit(GuiCtrlObj, Info) {
	global G_Combo, G_IniConfig, G_ErrText, G_LastQuery
	q := Trim(G_Combo ? G_Combo.Text : "")
	if (q = "" || q = G_ErrText || IsPathItemText(q))
		return

	p := ExpandPath(q)
	if (FileExist(p) != "") {
		JumpToPath(p)
		G_IniConfig.AddSearchHistory(q)
		RefreshComboItems(False)
		G_Combo.Text := ""
		return
	}
	if (IsPathLike(q)) {
		ShowSearchError(INVALID_TEXT)
		return
	}

	res := Ev3_Search(q)
	if (!IsObject(res)) {                       ; 不可用：DLL 缺失 或 Everything 未运行
		ShowSearchError("Everything 未运行")
		return
	}
	if (!res.Length) {
		ShowSearchError("无匹配结果")
		return
	}
	G_LastQuery := q
	ShowResultList(res, True)
	ToolTip("找到 " . res.Length . " 项")
	SetTimer(() => ToolTip(), -1800)
}

; 实时搜索（边打边搜）：输入每次变化都重置一次性定时器 → 停手 LIVE_DELAY 毫秒才查
LiveSearch() {
	global G_Combo, G_DlgHwnd, G_ErrText, G_LastQuery
	if (!G_Combo || !DllCall("IsWindow", "Ptr", G_DlgHwnd))
		return
	q := Trim(G_Combo.Text)
	if (q = "" || q = G_ErrText) {
		G_LastQuery := ""
		HideResultList()
		return
	}
	; 与上次相同、或属于路径类（交给回车处理）→ 不查
	if (q = G_LastQuery || IsPathItemText(q) || IsPathLike(q))
		return
	G_LastQuery := q

	res := Ev3_Search(q)
	if (!IsObject(res) || !res.Length) {
		HideResultList()
		return
	}
	ShowResultList(res, False)      ; 不抢焦点，避免打断打字
}

; 搜索框红字提示 + 系统错误提示音，约 1.8 秒后自动恢复
ShowSearchError(msg) {
	global G_Combo, G_ErrText
	if (!G_Combo)
		return
	G_ErrText := msg
	try G_Combo.SetFont("cRed")
	G_Combo.Text := msg
	DllCall("MessageBeep", "UInt", 0x10)                ; MB_ICONHAND：系统"错误"音
	SetTimer(ClearSearchError, -1800)
}

ClearSearchError() {
	global G_Combo, G_ErrText
	if (!G_Combo)
		return
	try G_Combo.SetFont("cDefault")
	if (G_ErrText != "" && G_Combo.Text = G_ErrText)
		G_Combo.Text := ""
	G_ErrText := ""
}

; ============================================================================
; 6 结果列表：独立 ListView 窗口（贴工具条下方；下方放不下则翻到对话框上方）
; ============================================================================

CreateListWnd() {
	global G_ListGui, G_List
	if (G_ListGui)
		return
	g := Gui("+ToolWindow -Caption +AlwaysOnTop", "QuickSwitch Results")
	g.MarginX := 0
	g.MarginY := 0
	g.BackColor := "FFFFFF"
	G_List := g.AddListView("x0 y0 w600 h200 Grid", ["名称", "完整路径"])
	G_List.OnEvent("DoubleClick", OnListPick)
	; 隐藏默认按钮：焦点在列表里时按回车也能跳转
	_hidden := g.AddButton("Default Hidden w1 h1", "")
	_hidden.OnEvent("Click", OnListEnter)
	g.OnEvent("Escape", OnListClose)
	g.OnEvent("Close", OnListClose)
	G_ListGui := g
}

; 用结果刷新并显示列表。
; activate=True 抢焦点（回车触发时用）；实时搜索必须 False，否则会打断打字。
; 批量填充期间关重绘(-Redraw)，否则逐行刷新会卡顿。
ShowResultList(paths, activate := False) {
	global G_List, G_ListGui, G_ListHwnd, G_ListItems, G_ListH
	CreateListWnd()
	G_ListItems := paths
	G_List.Opt("-Redraw")
	G_List.Delete()
	for p in paths {
		SplitPath(p, &_name)
		G_List.Add("", _name, p)
	}
	_rows := (paths.Length > 12) ? 12 : paths.Length
	G_ListH := _rows * 22 + 30
	if (G_ListH < 66)
		G_ListH := 66
	if (!ListGeom(&_x, &_y, &_w, &_h)) {
		G_List.Opt("+Redraw")
		return
	}
	G_List.Move(0, 0, _w, _h)
	G_List.ModifyCol(1, Round(_w * 0.38))
	G_List.ModifyCol(2, Round(_w * 0.62) - 22)
	G_List.Opt("+Redraw")
	G_ListGui.Show(Format("NA x{} y{} w{} h{}", _x, _y, _w, _h))
	G_ListHwnd := G_ListGui.Hwnd
	if (activate) {
		G_List.Modify(1, "Select Focus Vis")
		try WinActivate(G_ListHwnd)
		G_List.Focus()
	}
}

HideResultList() {
	global G_ListGui, G_ListHwnd
	G_ListHwnd := 0
	if (G_ListGui)
		try G_ListGui.Hide()
}

; 计算列表几何（屏幕坐标）。纯 DllCall/NumGet，可被 Fast 回调调用。
ListGeom(&x, &y, &w, &h) {
	global G_DlgHwnd, G_ListH, WK_L, WK_T, WK_R, WK_B
	static _r := Buffer(16)
	x := 0, y := 0, w := 0, h := G_ListH
	if (!DllCall("GetWindowRect", "Ptr", G_DlgHwnd, "Ptr", _r))
		return False
	_dl := NumGet(_r, 0, "Int")
	_dt := NumGet(_r, 4, "Int")
	_db := NumGet(_r, 12, "Int")
	w := NumGet(_r, 8, "Int") - _dl
	if (w < 300)
		w := 300
	x := _dl
	y := _db + BAR_H                        ; 优先贴工具条下方
	if (y + h > WK_B) {                     ; 下方放不下 → 翻到对话框上方
		y := _dt - h
		if (y < WK_T)
			y := WK_T
	}
	if (x + w > WK_R)
		x := WK_R - w
	if (x < WK_L)
		x := WK_L
	return True
}

; 跟随对话框移动（只移动，不改尺寸/层位 → Fast 回调安全）
RepositionResultList() {
	global G_ListHwnd
	if (!G_ListHwnd)
		return
	if (!ListGeom(&_x, &_y, &_w, &_h))
		return
	DllCall("SetWindowPos", "Ptr", G_ListHwnd, "Ptr", 0, "Int", _x, "Int", _y, "Int", 0, "Int", 0, "UInt", SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW | SWP_NOSENDCHANGING, "Int")
}

OnListPick(LV, Row) {
	JumpListItem(Row)
}

OnListEnter(GuiCtrlObj, Info) {
	global G_List
	JumpListItem(G_List.GetNext())
}

OnListClose(GuiObj) {
	HideResultList()
}

; 选中第 row 行 → 跳转
JumpListItem(row) {
	global G_ListItems
	if (row >= 1 && row <= G_ListItems.Length) {
		p := G_ListItems[row]
		HideResultList()
		JumpToPath(p)
	}
}

; ============================================================================
; 7 路径与对话框
; ============================================================================

; 跳转到指定路径：目录直达；文件则跳到所在目录并把文件名填进"文件名"框(≈选中该文件)
JumpToPath(p) {
	global G_DlgHwnd, G_Jumping
	if (G_Jumping || !DllCall("IsWindow", "Ptr", G_DlgHwnd))
		return False
	attr := FileExist(p)
	if (attr = "")
		return False
	G_Jumping := True
	try {
		if (InStr(attr, "D")) {
			FeedDialog(G_DlgHwnd, p, True)
		} else {
			SplitPath(p, &_name, &_dir)
			FeedDialog(G_DlgHwnd, _dir, True)
			Sleep 150
			SetDlgFileName(G_DlgHwnd, _name)
		}
	} catch {
		G_Jumping := False
		return False
	}
	G_Jumping := False
	return True
}

; 下拉项文本是否是"最近文件夹路径"项（带前缀）
IsPathItemText(t) {
	return SubStr(t, 1, StrLen(COMBO_PATH_PREFIX)) = COMBO_PATH_PREFIX
}

; 去掉下拉项前缀，取出真实路径
StripItemPrefix(t) {
	if (IsPathItemText(t))
		return Trim(SubStr(t, StrLen(COMBO_PATH_PREFIX) + 1))
	return t
}

; 看起来像路径（含 \ 或 :）——用于区分"打错的路径"与"搜索关键词"
IsPathLike(s) {
	return (InStr(s, "\") || InStr(s, ":")) ? True : False
}

; 去掉首尾成对引号，并展开 %环境变量%
ExpandPath(s) {
	s := Trim(s)
	len := StrLen(s)
	if (len >= 2 && ((SubStr(s, 1, 1) = '"' && SubStr(s, len) = '"') || (SubStr(s, 1, 1) = "'" && SubStr(s, len) = "'")))
		s := SubStr(s, 2, len - 2)
	if (InStr(s, "%")) {
		buf := Buffer(65536)
		if (DllCall("ExpandEnvironmentStringsW", "WStr", s, "Ptr", buf, "UInt", 32767))
			s := StrGet(buf)
	}
	return s
}

; 判断目录是否存在
ValidFolder(path) {
	if (path = "")
		return False
	return InStr(FileExist(path), "D") ? True : False
}

; 把文件名填进公共对话框的"文件名"输入框(Edit1)。
; ⚠ 导航后对话框会重建控件，刚取到的句柄可能随即失效（"target control not found"）
;   → 重试 + 校验，且任何异常都不外抛。
SetDlgFileName(_WinID, name, retries := 6) {
	Loop retries {
		try {
			_edit := ControlGetHwnd("Edit1", _WinID)
			if (_edit) {
				ControlSetText(name, "ahk_id " _edit)
				Sleep 25
				if (InStr(ControlGetText("Edit1", _WinID), name))   ; 导航若未结束，名字可能被重置
					return True
			}
		} catch {
			; 控件正在重建：忽略，稍后重试
		}
		Sleep 60
	}
	return False
}

; 取相邻主窗口里当前打开的文件夹路径（资源管理器 / TC / XYPlorer / Tablacus 等）
GetZWindowPath(WinID) {
	ZFolder := ""

	; 取"本窗口之后出现的第一个顶层窗口"——通常是对话框的宿主文件管理器
	root := DllCall("GetAncestor", "Ptr", WinID, "Uint", 3)     ; GA_ROOTOWNER
	WinIDs := WinGetList()
	found_idx := -1
	Loop WinIDs.Length {
		if (WinIDs[A_Index] = root)
			found_idx := A_Index
	}
	nextWinID := WinIDs[found_idx + 1]
	if (!nextWinID)
		return ZFolder

	Switch WinGetClass(nextWinID), True {
		Case "CabinetWClass":                                   ; 文件资源管理器
			activeTab := ControlGetHwnd("ShellTabWindowClass1", nextWinID)
			shellApp := ComObject("Shell.Application")
			For ComWin in shellApp.Windows {
				try
					_checkID := ComWin.hwnd
				catch
					continue
				if (nextWinID != _checkID)
					continue
				if (activeTab) {
					static IID_IShellBrowser := "{000214E2-0000-0000-C000-000000000046}"
					shellBrowser := ComObjQuery(ComWin, IID_IShellBrowser, IID_IShellBrowser)
					thisTab := 0
					ComCall(3, shellBrowser, "Ptr*", &thisTab)
					if (thisTab != activeTab)
						continue
				}
				ZFolder := ComWin.Document.Folder.Self.Path
				Break
			}
		Case "TablacusExplorer":
			shellApp := ComObject("Shell.Application")
			For ComWin in shellApp.Windows {
				SplitPath(ComWin.FullName, &ProcName)
				if !(ProcName ~= "i)TE(32|64)\.exe")
					continue
				if !(ParentWin := ComWin.Document.parentWindow)
					continue
				if (ParentWin.te.hwnd = nextWinID)
					ZFolder := ParentWin.te.Ctrl(0).FolderItem.Path
			}
		Case "TTOTAL_CMD":                                      ; Total Commander
		Case "dopus.lister":                                    ; Directory Opus
		Case "ThunderRT6FormDC":                                ; XYPlorer
	}
	return ZFolder
}

; 返回对话框样式（空 = 不认识的对话框）
GetFileDialogType(WinID) {
	try
		_ctrls := WinGetControls(WinID)
	catch
		return ""

	_has := Map()
	for _ctrl in _ctrls
		_has[_ctrl] := True
	if (!(_has.Has("ToolbarWindow321") && _has.Has("Edit1") && _has.Has("ComboBoxEx321")))
		return ""
	if (_has.Has("DirectUIHWND1"))
		return "GENERAL"                                        ; 现代公共对话框
	if (_has.Has("SysListView321"))
		return "SYSLISTVIEW"                                    ; Explorer 样式
	return ""
}

; 让公共对话框跳到指定目录
;   _safe=True：用户主动跳转（双击结果/输入路径）时用，回车后把焦点收回文件列表兜底
FeedDialog(_WinID, _NewPath, _safe := False) {
	switch G_FDType {
		case "GENERAL", "SYSLISTVIEW":
			FeedDialog_GENERAL(_WinID, _NewPath, _safe)
	}
}

FeedDialog_GENERAL(_WinID, _NewPath, _safe := False) {
	; 文件名输入框：兼作"这是不是我们能处理的对话框"的检查
	_edit := 0
	try
		_edit := ControlGetHwnd("Edit1", _WinID)
	catch
		_edit := 0
	if (!_edit) {
		MsgBox("This type of dialog can not be handled (yet).`nPlease report it!1")
		return
	}

	Sleep(100) ; 等控件稳定：某些情况下控件可能还没完全创建
	_addr_edit := FindChildByXPath("WorkerW/ReBarWindow32/Address Band Root/msctls_progress32/ComboBoxEx32/ComboBox/Edit", _WinID)
	if (!_addr_edit) {
		MsgBox("This type of dialog can not be handled (yet).`nPlease report it!2")
		return
	}
	_NewPath := RTrim(_NewPath, "\") . "\"

	try WinActivate(_WinID)

	; —— 关键：让焦点离开"文件夹/文件名"输入框(Edit1) ——
	; 回车若被对话框的消息循环吃掉，会落到**当前焦点控件**上；而文件夹对话框里
	; Edit1 就是"文件夹"框，回车 = 确定 → 对话框被直接确认（文件对话框的 Edit1 是
	; 文件名框、通常为空，所以只是无事发生，看起来"正常"）。
	; 实测 ControlFocus 对地址栏无效、对文件列表有效，所以聚焦文件列表并校验。
	_focusOK := FocusFileList(_WinID)
	Sleep(50)

	; —— 导航：写地址栏 + 给地址栏发回车 ——
	; 实测（自建"选择文件夹"对话框复现，用 WinGetText 判定是否真导航）：
	;   * ControlFocus 无法把焦点交给地址栏；Alt+D 会重置地址栏文本 → 两者都让导航失效；
	;   * 唯一有效的是下面这段：回车由地址栏自己消费，只做导航。
	try {
		ControlSetText(_NewPath, _addr_edit, _WinID)
		Sleep(50)
		if (_safe && !_focusOK)
			return              ; 焦点没能离开 Edit1 → 宁可不导航，也不冒"被确定"的风险
		ControlSend("{Enter}", _addr_edit, _WinID)
		Sleep(120)
		FocusFileList(_WinID, 2)    ; 回车后也把焦点收回列表，兜住滞留的回车
	}
	Sleep(10)
}

; 把键盘焦点移到对话框的文件列表上，并**校验**确实移到了（返回 True = 焦点不在 Edit1 上）。
; 为什么要校验：回车一旦落到 Edit1（文件夹框）就等于"确定"。
FocusFileList(_WinID, tries := 4) {
	global G_FDType
	Loop tries {
		try
			ControlFocus("DirectUIHWND2", "ahk_id " _WinID)
		catch
			try ControlFocus("DirectUIHWND1", "ahk_id " _WinID)
		Sleep 40
		if (GetFocusedCtrlHwnd(_WinID) = FileListHwnd(_WinID))
			return True
		try WinActivate("ahk_id " _WinID)
		Sleep 60
	}
	return False
}

; 文件列表控件的 hwnd（现代对话框是 DirectUIHWND2/1，老样式是 SysListView321）
FileListHwnd(_WinID) {
	for cls in ["DirectUIHWND2", "DirectUIHWND1", "SysListView321"] {
		try {
			hwnd := ControlGetHwnd(cls, "ahk_id " _WinID)
			if (hwnd)
				return hwnd
		}
	}
	return 0
}

; 取对话框里当前拥有键盘焦点的控件 hwnd（取不到返回 0）
GetFocusedCtrlHwnd(_WinID) {
	try {
		cls := ControlGetFocus("ahk_id " _WinID)
		if (cls)
			return ControlGetHwnd(cls, _WinID)
	}
	return 0
}

; 按 "类名/类名/..." 逐层查找子控件，返回其 hwnd（找不到返回 0）
FindChildByXPath(cls_xpath, hwnd) {
	fpos := InStr(cls_xpath, "/")
	if (!fpos)
		return _FindChildByClass(cls_xpath, hwnd)
	return _FindChildByClass(SubStr(cls_xpath, 1, fpos - 1), hwnd, SubStr(cls_xpath, fpos + 1))
}
_FindChildByClass(cls, hwnd, rest := "") {
	GW_CHILD := 5
	GW_HWNDFIRST := 0
	GW_HWNDNEXT := 2
	_child := DllCall("GetWindow", "Ptr", hwnd, "Uint", GW_CHILD | GW_HWNDFIRST)
	while (_child) {
		if (WinGetClass(_child) = cls)
			return rest = "" ? _child : FindChildByXPath(rest, _child)
		_child := DllCall("GetWindow", "Ptr", _child, "Uint", GW_HWNDNEXT)
	}
	return 0
}

; ============================================================================
; 8 配置（config.ini）
; ============================================================================
Class IniConfig {
	_filename := ""

	__New(filename) {
		this._filename := filename
	}

	; —— [Dialogs] 各对话框的自动切换开关 ——
	ReadDialogNode(DialogStr) {
		return IniRead(this._filename, "Dialogs", DialogStr, "")
	}
	WriteDialogNode(DialogStr, Enable := 1) {
		IniWrite(Format("{}={}", DialogStr, Enable), this._filename, "Dialogs")
	}

	; —— 各对话框的"最近路径"列表（写在以对话框标识命名的节里）——
	ReadDialogPaths(DialogStr) {
		return this._Split(IniRead(this._filename, DialogStr, "paths", ""))
	}
	WriteDialogPaths(DialogStr, path) {
		arr := this.ReadDialogPaths(DialogStr)
		for p in arr
			if (p = path)
				return                                  ; 已存在，不重复记
		if (arr.Length > MAX_DIALOG_PATHS)
			arr.Pop()
		arr.InsertAt(1, path)
		IniWrite(this._Join(arr), this._filename, DialogStr, "paths")
	}

	; —— [Settings] 通用设置 ——
	ReadSetting(Key, Default := "") {
		return IniRead(this._filename, "Settings", Key, Default)
	}
	WriteSetting(Key, Value) {
		IniWrite(Value, this._filename, "Settings", Key)
	}

	; —— [Search] 输入历史 ——
	ReadSearchHistory() {
		return this._Split(IniRead(this._filename, "Search", "history", ""))
	}
	WriteSearchHistory(arr) {
		IniWrite(this._Join(arr), this._filename, "Search", "history")
	}
	; 新增一条：去重 → 置顶 → 按上限截断
	AddSearchHistory(q) {
		q := Trim(q)
		if (q = "")
			return
		arr := this.ReadSearchHistory()
		Loop arr.Length {
			if (arr[A_Index] = q) {
				arr.RemoveAt(A_Index)
				Break
			}
		}
		arr.InsertAt(1, q)
		max := this.HistoryMax()
		while (arr.Length > max)
			arr.Pop()
		this.WriteSearchHistory(arr)
	}
	ClearSearchHistory() {
		this.WriteSearchHistory([])
	}
	HistoryMax() {
		n := this.ReadSetting("history_max", HISTORY_MAX_DEF) + 0
		return (n < 1) ? HISTORY_MAX_DEF : n
	}

	; —— 内部：数组 <-> "a|b|c" ——
	_Join(arr) {
		s := ""
		for v in arr
			s := (s = "") ? v : s . "|" . v
		return s
	}
	_Split(s) {
		arr := []
		if (s != "")
			for v in StrSplit(s, "|")
				if (v != "")
					arr.Push(v)
		return arr
	}
}

; ============================================================================
; 9 首选项
; ============================================================================
; 已实现：启用工具条 / 随系统启动 / 新对话框默认自动切换 / 历史条数 / 清空历史 / 打开配置文件
; 规划中：工具条外观、自动切换延迟与进程黑名单、搜索范围与条数、
;         支持的文件管理器开关、全局热键、配置导入导出/重置、主题
ShowPreferences() {
	global G_PrefGui, G_PrefBar, G_PrefAutoRun, G_PrefAutoSw, G_PrefMax
	global G_IniConfig, G_bEnableBar, G_bDefaultAutoSwitch
	if (!G_PrefGui) {
		g := Gui("+AlwaysOnTop", "QuickSwitch 首选项")
		g.MarginX := 14
		g.MarginY := 12
		g.SetFont("s9")

		g.AddText("xm w320", "通用")
		G_PrefBar := g.AddCheckBox("xm+14 y+6", "启用底部搜索工具条")
		G_PrefAutoRun := g.AddCheckBox("xm+14 y+4", "随系统启动")

		g.AddText("xm y+12 w320", "自动切换")
		G_PrefAutoSw := g.AddCheckBox("xm+14 y+6", "新对话框默认开启自动切换")

		g.AddText("xm y+12 w320", "搜索与历史")
		g.AddText("xm+14 y+8", "最多保留历史条数：")
		G_PrefMax := g.AddEdit("x+6 yp-2 w60 Number")

		g.AddText("xm y+14 w320", "其他")
		g.AddButton("xm+14 y+6 w120", "打开配置文件").OnEvent("Click", OnPrefsOpenIni)
		g.AddButton("x+10 w120", "清空搜索历史").OnEvent("Click", OnPrefsClearHistory)

		g.AddButton("xm+90 y+20 w80 Default", "确定").OnEvent("Click", OnPrefsOK)
		g.AddButton("x+10 w80", "取消").OnEvent("Click", OnPrefsCancel)

		G_PrefGui := g
	}
	; 每次打开都按当前设置刷新控件
	G_PrefBar.Value := G_bEnableBar ? 1 : 0
	G_PrefAutoSw.Value := G_bDefaultAutoSwitch ? 1 : 0
	G_PrefMax.Value := G_IniConfig.ReadSetting("history_max", HISTORY_MAX_DEF)
	G_PrefAutoRun.Value := IsAutoRun() ? 1 : 0
	G_PrefGui.Show()
}

OnPrefsOK(GuiCtrlObj, Info) {
	global G_PrefGui, G_PrefBar, G_PrefAutoRun, G_PrefAutoSw, G_PrefMax
	global G_IniConfig, G_bEnableBar, G_bDefaultAutoSwitch
	G_bEnableBar := (G_PrefBar.Value = 1)
	G_bDefaultAutoSwitch := (G_PrefAutoSw.Value = 1)
	n := G_PrefMax.Value + 0
	if (n < 1)
		n := HISTORY_MAX_DEF
	G_PrefMax.Value := n
	G_IniConfig.WriteSetting("enable_bar", G_bEnableBar ? 1 : 0)
	G_IniConfig.WriteSetting("default_autoswitch", G_bDefaultAutoSwitch ? 1 : 0)
	G_IniConfig.WriteSetting("history_max", n)
	SetAutoRun(G_PrefAutoRun.Value = 1)
	G_PrefGui.Hide()
}

OnPrefsCancel(GuiCtrlObj, Info) {
	global G_PrefGui
	G_PrefGui.Hide()
}

OnPrefsClearHistory(GuiCtrlObj, Info) {
	global G_IniConfig
	G_IniConfig.ClearSearchHistory()
	RefreshComboItems(False)
	ToolTip("已清空搜索历史")
	SetTimer(() => ToolTip(), -1200)
}

OnPrefsOpenIni(GuiCtrlObj, Info) {
	Run('notepad.exe "' . A_ScriptDir . '\config.ini"')
}

; 是否已设置开机自启（HKCU\...\Run）
IsAutoRun() {
	try
		return RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Run", "QuickSwitch") != ""
	catch
		return False
}

SetAutoRun(enable) {
	key := "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"
	if (enable) {
		cmd := A_IsCompiled ? Format('"{1}"', A_ScriptFullPath) : Format('"{1}" "{2}"', A_AhkPath, A_ScriptFullPath)
		RegWrite(cmd, "REG_SZ", key, "QuickSwitch")
	} else {
		try RegDelete(key, "QuickSwitch")
	}
}

; ============================================================================
; 10 杂项工具
; ============================================================================

; 隐藏指定窗口
WinHideById(hwnd) {
	DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", 0, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", SWP_HIDEWINDOW | SWP_NOACTIVATE, "Int")
}
