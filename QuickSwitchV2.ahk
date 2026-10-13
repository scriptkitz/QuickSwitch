#Requires AutoHotkey >=2.0
#Singleinstance Force

#Include "Sqlite.ahk"

SetWinDelay -1
SetControlDelay -1

SetWorkingDir(A_ScriptDir)

If (VerCompare(A_OSVersion, "10.0.0") <= 0) {
	MsgBox(A_OSVersion " is not supported.")
	ExitApp
}

G_IniConfig := IniConfig("config.ini")

Hotkey("^W", ShowMenu, "Off")

SWP_NOSIZE := 0x1
SWP_NOACTIVATE := 0x10
SWP_DRAWFRAME  := 0x20
SWP_SHOWWINDOW := 0x40
SWP_HIDEWINDOW := 0x80
SWP_ASYNCWINDOWPOS := 0x4000

WS_POPUP   := 0x80000000
WS_CHILD   := 0x40000000
WS_VISIBLE := 0x10000000

SWP_NOZORDER := 0x4
SWP_NOSENDCHANGING := 0x400
EVENT_OBJECT_LOCATIONCHANGE := 0x800B
OBJID_WINDOW := 0

WM_CHANGEUISTATE := 0x0127
WM_UPDATEUISTATE := 0x0128

; 底部工具条几何参数：[ ComboBox ] [ ⚙ ] 同一父窗口并排，贴在对话框下边缘外侧
BAR_H := 30
BAR_BTN_W := 32
BAR_GAP := 4
COMBO_PATH_PREFIX := "📁 "   ; 下拉里"最近文件夹路径"项的前缀，用于区分选中项类型
INVALID_TEXT := "无效路径"    ; 输入非法路径时在搜索框里显示的红色提示

G_DlgHwnd := 0
G_LastSetDlg := 0 ; 最后一次自动切换路径的对话框
G_Menu := Menu()
G_Combo := 0             ; ComboBox 控件对象(用于增删下拉项)
G_ComboHwnd := 0
G_ComboPaths := []       ; 下拉里"路径项"的路径数组(与项索引一一对应)
G_BarBtnHwnd := 0
; 预计算的布局常量（在 CreateBarWnd 里填；重定位时不再重复测量）
G_ComboWinH := 0
G_ComboY := 3
G_BtnY := 3
G_BarChildW := -1   ; 子控件上次按多宽排过版
G_BarGui := CreateBarWnd()
G_BarHwnd := G_BarGui.Hwnd
G_SearchText := ""
; 设置项
G_bEnableBar := G_IniConfig.ReadSetting("enable_bar", 1) ? True : False
G_bDefaultAutoSwitch := G_IniConfig.ReadSetting("default_autoswitch", 1) ? True : False
G_PrefGui := 0
G_PrefBar := 0
G_PrefAutoRun := 0
G_PrefAutoSw := 0
G_PrefMax := 0
; 窗口事件钩子：不全局注册，而是在检测到对话框时按"对话框所在线程"注册，
; 只接收该线程的事件，大幅减少无关进程的唤醒开销（WINEVENT_OUTOFCONTEXT = flags 0）。
G_WinEventProc := CallbackCreate(WinEventProc, "Fast", 7)
G_WinEventHook := 0
_G_DialogFinger := ""
_G_FDType := ""
_G_bAutoSwitch := True

Loop {
	WinID := WinWaitActive("ahk_class #32770")
	_G_FDType := GetFileDialogType(WinID)
	if (_G_FDType) {
		G_DlgHwnd := WinID
		windows_title := WinGetTitle(WinID)
		windows_exe := WinGetProcessName(WinID)
		_G_DialogFinger := windows_exe . "___" . windows_title
		ShowBar(WinID)
		; 激活的窗口在失去激活前只能自动设置一次路径。
		if (G_LastSetDlg = G_DlgHwnd) { 
			Continue
		}
		_bAuto := G_IniConfig.ReadDialogNode(_G_DialogFinger)
		if (_bAuto == "") {
			_bAuto := G_bDefaultAutoSwitch ? 1 : 0
			G_IniConfig.WriteDialogNode(_G_DialogFinger, _bAuto)
		}
		_G_bAutoSwitch := _bAuto ? True : False
		FolderPath := GetZWindowPath(WinID)
		if (_G_bAutoSwitch && ValidFolder(FolderPath)) {
			FeedDialog(WinID, FolderPath)
		}

		WaitDialogSessionEnd(WinID)
		G_IniConfig.WriteDialogPaths(_G_DialogFinger, FolderPath)
		HideBar()
		G_LastSetDlg := 0
	}
}
ExitApp(0)

MAKELONG(LOWORD,HIWORD,Hex:=0){
    BITS:=0x10,WORD:=0xFFFF
    return (!Hex)?((HIWORD<<BITS)|(LOWORD&WORD)):Format("{1:#x}",((HIWORD<<BITS)|(LOWORD&WORD)))
}

OnBtnSetting(Obj, info) {
	global G_BarHwnd
	AUTOSWITCH := "自动切换"
	PREFS := "首选项"
	OnMenuPath(_name, _pos, _menu) {
		if (AUTOSWITCH = _name) {
			_menu.ToggleCheck(_name)
			G_IniConfig.WriteDialogNode(_G_DialogFinger, !_G_bAutoSwitch)
		} else if (PREFS = _name) {
			; 用一次性定时器推迟到菜单关闭后再弹窗
			SetTimer(ShowPreferences, -10)
		}
	}
	MouseGetPos &x, &y
	G_Menu.Delete()
	G_Menu.Add(AUTOSWITCH, OnMenuPath)
	_G_bAutoSwitch ? G_Menu.Check(AUTOSWITCH) : G_Menu.UnCheck(AUTOSWITCH)
	G_Menu.Add()   ; 分隔线
	G_Menu.Add(PREFS, OnMenuPath)
	ControlGetPos &cx, &cy, &cw, &ch, Obj

	; 菜单挂到工具条窗口下，显示在 ⚙ 按钮正下方
	_parent := DllCall("GetParent", "Ptr", G_Menu.Handle)
	if (_parent != G_BarHwnd) {
		DllCall("SetParent", "Ptr", G_Menu.Handle, "Ptr", G_BarHwnd)
	}
	G_Menu.Show(cx, cy+ch)
}

; ===== 底部工具条（单一父窗口）：[ ComboBox ] [ ⚙ ]，跟随系统文件对话框显示与移动 =====
CreateBarWnd() {
	global G_ComboHwnd, G_BarBtnHwnd, BAR_H, BAR_BTN_W
	global G_ComboWinH, G_ComboY, G_BtnY
	global G_Combo
	_gui := Gui("+ToolWindow -Caption +AlwaysOnTop")
	_gui.MarginX := 4
	_gui.MarginY := 3
	_gui.BackColor := "F2F2F2"
	; ComboBox：以后用 Add()/Delete() 放输入历史，可下拉选择（h200 为下拉列表高度）
	_combo := _gui.AddComboBox("w280 h200", [])
	G_Combo := _combo
	G_ComboHwnd := _combo.Hwnd
	_combo.OnEvent("Change", OnSearchInput)
	; ⚙ 按钮（位置/尺寸由 RepositionBar 统一布局；未来承载菜单/配置）
	_btn := _gui.AddButton("x300 y3 w32 h24", "⚙")
	G_BarBtnHwnd := _btn.Hwnd
	_btn.OnEvent("Click", OnBtnSetting)
	; 隐藏默认按钮：捕获 Enter（ComboBox 没有 KeyDown 事件）
	_hbtn := _gui.AddButton("Default Hidden w1 h1", "")
	_hbtn.OnEvent("Click", OnSearchSubmit)

	; —— 预计算布局常量：只测一次，之后每次重定位不再 GetWindowRect/SendMessage ——
	_r := Buffer(16)
	DllCall("GetWindowRect", "Ptr", G_ComboHwnd, "Ptr", _r)
	G_ComboWinH := NumGet(_r, 12, "Int") - NumGet(_r, 4, "Int")   ; ComboBox 窗口总高(含下拉区)
	_vh := DllCall("SendMessage", "Ptr", G_ComboHwnd, "UInt", 0x0154, "Ptr", -1, "Ptr", 0, "Ptr")  ; 选择区可见高度
	if (!(_vh >= 16 && _vh <= BAR_H - 4))
		_vh := 22
	G_ComboY := (BAR_H - _vh) // 2
	G_BtnY := (BAR_H - (BAR_H - 6)) // 2
	return _gui
}

; 显示底部工具条，并把位置钩子限定到该对话框所在线程。
ShowBar(_WinID) {
	global G_BarGui, G_BarHwnd, G_BarChildW, G_bEnableBar
	if (!G_BarGui || !G_bEnableBar)
		return
	G_BarHwnd := G_BarGui.Hwnd
	G_BarChildW := -1          ; 强制重新排布子控件
	RefreshComboItems(True)    ; 按当前对话框重建下拉：最近文件夹路径 + 搜索关键词历史
	HookDialogThread(_WinID)
	RepositionBar(_WinID)
}

HideBar() {
	global G_BarHwnd
	UnhookDialogThread()
	if (G_BarHwnd && DllCall("IsWindow", "Ptr", G_BarHwnd))
		DllCall("SetWindowPos", "Ptr", G_BarHwnd, "Ptr", 0, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", SWP_HIDEWINDOW | SWP_NOACTIVATE, "Int")
	G_BarHwnd := 0
}

; 重建 ComboBox 下拉项：先"最近文件夹路径"(📁 前缀)，再"输入历史"。
; clearText=True 时同时清空输入框（新对话框会话开始时用）。
RefreshComboItems(clearText := False) {
	global G_Combo, G_ComboPaths, G_IniConfig, _G_DialogFinger
	if (!G_Combo)
		return
	paths := G_IniConfig.ReadDialogPaths(_G_DialogFinger)
	kws := G_IniConfig.ReadSearchHistory()
	G_ComboPaths := paths
	items := []
	for p in paths
		items.Push(COMBO_PATH_PREFIX . p)
	for k in kws
		items.Push(k)
	G_Combo.Delete()
	if (items.Length)
		G_Combo.Add(items)
	if (clearText)
		G_Combo.Text := ""
}

; 只监听指定窗口所属线程的位置/尺寸事件（跨进程用 out-of-context 钩子）。
HookDialogThread(winID) {
	global G_WinEventHook, G_WinEventProc, EVENT_OBJECT_LOCATIONCHANGE
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

; 把工具条贴到对话框下边缘外侧（屏幕坐标），并排好内部 [ ComboBox | ⚙ ]。
; 可被 Fast 回调直接调用：只用 DllCall/Buffer/NumGet。
RepositionBar(_WinID) {
	global G_BarHwnd, G_ComboHwnd, G_BarBtnHwnd, BAR_H, BAR_BTN_W, BAR_GAP
	global G_ComboWinH, G_ComboY, G_BtnY, G_BarChildW
	static _r := Buffer(16)
	if (!G_BarHwnd)
		return
	; 对话框已销毁时 GetWindowRect 失败 → 直接返回，省掉一次 IsWindow
	if (!DllCall("GetWindowRect", "Ptr", _WinID, "Ptr", _r))
		return
	_l := NumGet(_r, 0, "Int")
	_rt := NumGet(_r, 8, "Int")
	_b := NumGet(_r, 12, "Int")
	_w := _rt - _l
	if (_w < 200)
		_w := 200
	; 工具条窗：贴在对话框下边缘外侧，宽度与对话框一致，用 SWP_NOZORDER 保持 AlwaysOnTop 层位
	DllCall("SetWindowPos", "Ptr", G_BarHwnd, "Ptr", 0, "Int", _l, "Int", _b, "Int", _w, "Int", BAR_H, "UInt", SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW | SWP_NOSENDCHANGING, "Int")
	; 子控件(子窗口)随父窗自动跟随移动，只有在宽度变化(缩放)时才需要重排
	if (_w != G_BarChildW) {
		G_BarChildW := _w
		DllCall("SetWindowPos", "Ptr", G_ComboHwnd, "Ptr", 0, "Int", 4, "Int", G_ComboY, "Int", _w - 8 - BAR_BTN_W - BAR_GAP, "Int", G_ComboWinH, "UInt", SWP_NOZORDER | SWP_NOACTIVATE, "Int")
		DllCall("SetWindowPos", "Ptr", G_BarBtnHwnd, "Ptr", 0, "Int", _w - 4 - BAR_BTN_W, "Int", G_BtnY, "Int", BAR_BTN_W, "Int", BAR_H - 6, "UInt", SWP_NOZORDER | SWP_NOACTIVATE, "Int")
	}
}

; 等待"对话框会话"结束。用 WinWaitNotActive（AHK 内置的前台等待，开销极低）代替 30ms 轮询；
; 仅当焦点落在工具条上时才短暂轮询，避免因工具条持有焦点而误判会话未结束。
WaitDialogSessionEnd(WinID) {
	global G_BarHwnd, G_PrefGui
	Loop {
		WinWaitNotActive(WinID)
		if (!DllCall("IsWindow", "Ptr", WinID))
			return
		if (G_BarHwnd && WinActive("ahk_id " G_BarHwnd)) {
			Sleep 30
			continue
		}
		; 首选项对话框打开时也算会话进行中，避免工具条被隐藏
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

; SetWinEventHook 回调：仅处理当前文件对话框自身(OBJID_WINDOW)的位置/尺寸变化
WinEventProc(hHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime) {
	global G_DlgHwnd, G_BarHwnd
	if (event == EVENT_OBJECT_LOCATIONCHANGE && idObject == OBJID_WINDOW && hwnd == G_DlgHwnd && G_BarHwnd) {
		RepositionBar(hwnd)
	}
}

; ComboBox Change：打字或从下拉里选择都会触发
OnSearchInput(GuiCtrlObj, Info) {
	global G_SearchText, G_DlgHwnd
	t := GuiCtrlObj.Text
	; 选中的是"最近文件夹路径"项 → 让文件对话框切换到该文件夹，并清空搜索框
	if (t != "" && SubStr(t, 1, StrLen(COMBO_PATH_PREFIX)) = COMBO_PATH_PREFIX) {
		p := Trim(SubStr(t, StrLen(COMBO_PATH_PREFIX) + 1))
		if (ValidFolder(p)) {
			FeedDialog(G_DlgHwnd, p)
			G_SearchText := ""
			GuiCtrlObj.Text := ""
			return
		}
	}
	G_SearchText := t
}

; 隐藏默认按钮的 Click：搜索框中按 Enter 时触发。
; 输入按"路径"处理：目录 → 直接跳转；文件 → 跳到所在目录并填好文件名；
; 无效 → 搜索框红字提示 + 系统错误提示音。
OnSearchSubmit(GuiCtrlObj, Info) {
	global G_SearchText, G_IniConfig, G_Combo, G_DlgHwnd
	q := Trim(G_SearchText)
	; 忽略：空输入、下拉里的"路径项"(选中时已处理)、我们自己的红字提示
	if (q = "" || q = INVALID_TEXT)
		return
	if (SubStr(q, 1, StrLen(COMBO_PATH_PREFIX)) = COMBO_PATH_PREFIX)
		return
	p := ExpandPath(q)
	attr := FileExist(p)
	if (attr = "") {
		ShowInvalidPath()
		return
	}
	if (InStr(attr, "D")) {
		FeedDialog(G_DlgHwnd, p)              ; 目录：直接跳转
	} else {
		SplitPath(p, &_name, &_dir)           ; 文件：先跳到所在目录
		FeedDialog(G_DlgHwnd, _dir)
		Sleep 120
		SetDlgFileName(G_DlgHwnd, _name)      ; 再把文件名填进"文件名"框（≈选中该文件）
	}
	; 记住这次输入并刷新下拉，然后清空搜索框
	G_IniConfig.AddSearchHistory(q)
	RefreshComboItems(False)
	if (G_Combo) {
		G_Combo.Text := ""
		G_SearchText := ""
	}
}

; 展开 %环境变量% 并去掉首尾成对引号
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

; 把文件名填进公共对话框的"文件名"输入框(Edit1)
SetDlgFileName(_WinID, name) {
	_edit := ControlGetHwnd("Edit1", _WinID)
	if (!_edit)
		return
	ControlSetText(name, "ahk_id " _edit)
}

; 无效路径：搜索框红字提示 + 系统错误提示音，1.5 秒后自动恢复
ShowInvalidPath() {
	global G_Combo, G_SearchText
	if (!G_Combo)
		return
	try G_Combo.SetFont("cRed")
	G_Combo.Text := INVALID_TEXT
	G_SearchText := ""
	DllCall("MessageBeep", "UInt", 0x10)   ; MB_ICONHAND：系统"错误"提示音
	SetTimer(ClearInvalidPath, -1500)
}

ClearInvalidPath() {
	global G_Combo
	if (!G_Combo)
		return
	try G_Combo.SetFont("cDefault")
	if (G_Combo.Text = INVALID_TEXT)
		G_Combo.Text := ""
}

; 严重目录是否正确
ValidFolder(_path_) {
	if(_path_ != "") {
		if (InStr(FileExist(_path_), "D"))
			return True
		Else
			return False
	}
	return False
}

; 获取下一个主窗口包含的文件夹路径
GetZWindowPath(WinID) {
	ZFolder := ""
	_GetNextWinID(WinID) {
		WinIDs := WinGetList()
		GA_PARENT := 1
		GA_ROOT := 2
		GA_ROOTOWNER := 3
		root := DllCall("GetAncestor", "Ptr", WinID, "Uint", GA_ROOTOWNER)
		found_idx := -1
		Loop WinIDs.Length {
			if (WinIDs[A_Index] = root) {
				found_idx := A_Index
			}
		}
		return WinIDs[found_idx + 1]
	}

	nextWinID := _GetNextWinID(WinID)
	found_class := WinGetClass(nextWinID)
	Switch found_class, True {
		Case "TTOTAL_CMD": ;	Total Commander
		Case "ThunderRT6FormDC": ;	XYPlorer
		Case "CabinetWClass": ;	File Explorer
			activeTab := ControlGetHwnd("ShellTabWindowClass1", nextWinID)
			shellApp := ComObject("Shell.Application")
			For ComWin in shellApp.Windows {
				try 
					_checkID := ComWin.hwnd
				catch
					continue
				if (nextWinID = _checkID) {
					if (activeTab)
					{
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
			}
		Case "TablacusExplorer": ;	
			shellApp := ComObject("Shell.Application")
			For (ComWin in shellApp.Windows)
			{
				SplitPath(ComWin.FullName, &ProcName)
				if !(ProcName ~= "i)TE(32|64)\.exe")
					continue
				if !(ParentWin := ComWin.Document.parentWindow)
					continue
				te := ParentWin.te
				if (te.hwnd = nextWinID)
				{
                    ZFolder := te.Ctrl(0).FolderItem.Path
				}
			}
		Case "dopus.lister": ;	Directory Opus
		Case "ThunderRT6FormDC": ;	XYPlorer
	}
	return ZFolder
}

ShowMenu(key) {
	MsgBox(A_OSVersion " is not supported.")
}

; 返回对话框样式
GetFileDialogType(WinID) {
	Try {
		_ctrls := WinGetControls(WinID)
	} catch {
		return ""
	}
	_SysListView321 := False
	_ToolbarWindow321 := False
	_DirectUIHWND1 := False
	_Edit1 := False
	_ComboBoxEx321 := False
	for _ctrl in _ctrls {
		if ("SysListView321" = _ctrl) {
			_SysListView321 := True
		} else if ("ToolbarWindow321" = _ctrl) {
			_ToolbarWindow321 := True
		} else if ("DirectUIHWND1" = _ctrl) {
			_DirectUIHWND1 := True
		} else if ("Edit1" = _ctrl) {
			_Edit1 := True
		} else if ("ComboBoxEx321" = _ctrl) {
			_ComboBoxEx321 := True
		}
	}
	if (_ToolbarWindow321 and _Edit1 And _ComboBoxEx321) {
		If ( _DirectUIHWND1 )
		{
			Return "GENERAL" ; 公共对话框
		}
		Else If ( _SysListView321 )
		{
			Return "SYSLISTVIEW" ; Exploer样式的对话框
		}
	}
	
	Return FALSE
}

FeedDialog(_WinID, _NewPath) {
	switch _G_FDType {
		case "GENERAL":
			FeedDialog_GENERAL(_WinID, _NewPath)
		case "SYSLISTVIEW":
			FeedDialog_GENERAL(_WinID, _NewPath)
	}
}

/**
 * 判断指定 hwnd 是否为目标类名窗口的后代（子级、孙级等）
 * @param {Ptr} childHwnd - 要检查的窗口句柄
 * @param {String} parentClass - 目标的父类名 (例如 "abc")
 * @returns {Ptr} - 如果匹配，返回找到的父窗口句柄；否则返回 0
 */
IsDescendantOfClass(childHwnd, parentClass) {
    currHwnd := childHwnd
    
    ; 循环获取父窗口，直到没有父窗口为止 (返回 0)
    ; 注意：GetParent 可以获取控件的父窗口，也可以获取子窗口的父窗口
    while (currHwnd := DllCall("GetParent", "Ptr", currHwnd, "Ptr")) {
        try {
            ; 获取当前父窗口的类名
            if (WinGetClass(currHwnd) = parentClass) {
                return currHwnd  ; 找到了匹配类名的祖先，返回其句柄
            }
        } catch {
            ; 某些系统窗口可能禁止访问，直接跳过
            continue
        }
    }
    
    return 0 ; 溯源到顶层也没找到
}

FindChildByXPath(cls_xpath, hwnd, bClassNN:=False) {
	GW_HWNDFIRST :=        0
	GW_HWNDLAST :=         1
	GW_HWNDNEXT :=         2
	GW_HWNDPREV :=         3
	GW_OWNER :=            4
	GW_CHILD :=            5
	cls_xpath := StrReplace(cls_xpath, "\", "/")
	fpos := InStr(cls_xpath, "/")
	_left_cls := ""
	if (!fpos) {
		_need_cls := cls_xpath
	} else {
		_need_cls := SubStr(cls_xpath, 1, fpos-1)
		_left_cls := SubStr(cls_xpath, fpos+1)
	}
	_child := DllCall("GetWindow", "Ptr", hwnd, "Uint", GW_CHILD|GW_HWNDFIRST)
	Loop {
		if(_child) {
			if (bClassNN)
				_cls := ControlGetClassNN(_child, hwnd)
			Else
				_cls := WinGetClass(_child)
			if (_need_cls = _cls) {
				if (_left_cls) {
					return FindChildByXPath(_left_cls, _child)
				} else {
					return _child
				}
			}
			_child := DllCall("GetWindow", "Ptr", _child, "Uint", GW_HWNDNEXT)
		}
	} Until !_child

	return 0
}

; 设置公共对话框的路径
FeedDialog_GENERAL(_WinID, _NewPath) {
	_AddressToolbar := ""
	; 获取文件名输入框控件
	_edit := ControlGetHwnd("Edit1", _WinID)
	If (!_edit) {
		MsgBox "This type of dialog can not be handled (yet).`nPlease report it!1"
		Return
	}
	Sleep(100) ; 等待控件稳定，某些情况下控件可能还未完全创建
	_addr_edit := FindChildByXPath("WorkerW/ReBarWindow32/Address Band Root/msctls_progress32/ComboBoxEx32/ComboBox/Edit", _WinID)
	If (!_addr_edit) {
		MsgBox "This type of dialog can not be handled (yet).`nPlease report it!2"
		Return
	}
	_NewPath := RTrim(_NewPath, "\") "\"
	
	WinActivate(_WinID)
	
	; 激活DirectUIHWND区域，确保地址栏的ComboBoxEx32能够接收到输入
	try {
		ControlFocus("DirectUIHWND2", "ahk_id " _WinID)
		; ControlClick("DirectUIHWND2", "ahk_id " _WinID, , "LEFT", 1, "NA x5 y5")
	} catch {
		ControlFocus("DirectUIHWND1", "ahk_id " _WinID)
		; ControlClick("DirectUIHWND1", "ahk_id " _WinID, , "LEFT", 1, "NA x5 y5")
	}
	Sleep(50)
	; Send("!d") ; 激活地址栏 Alt+D
	; Sleep(50)
	ControlSetText _NewPath, _addr_edit, _WinID
	Sleep(50)
	ControlSend("{Enter}", _addr_edit, _WinID)

	; ControlSend("{Enter down}", _addr_edit, _WinID)
	; WM_CHAR := 0x0102
	; PostMessage(WM_CHAR, 13, 0, _addr_edit, _WinID)
	; ControlSend("{Enter up}", _addr_edit, _WinID)
	; Sleep(50)
	
	Sleep(10)
}

; ===== 首选项对话框 =====
; 已实现：启用工具条 / 随系统启动 / 默认自动切换 / 历史条数 / 清空历史 / 打开配置文件
; 规划中（待接）：工具条外观(高度/按钮宽)、自动切换延迟、按进程黑名单、
;                 支持的文件管理器开关(Explorer/TC/XYPlorer/DOpus/Tablacus)、呼出快捷键、配置导入导出/重置
ShowPreferences() {
	global G_PrefGui, G_PrefBar, G_PrefAutoRun, G_PrefAutoSw, G_PrefMax
	global G_IniConfig, G_bEnableBar, G_bDefaultAutoSwitch
	if (!G_PrefGui) {
		_g := Gui("+AlwaysOnTop", "QuickSwitch 首选项")
		_g.MarginX := 14
		_g.MarginY := 12
		_g.SetFont("s9")

		_g.AddText("xm w320", "通用")
		G_PrefBar := _g.AddCheckBox("xm+14 y+6", "启用底部搜索工具条")
		G_PrefAutoRun := _g.AddCheckBox("xm+14 y+4", "随系统启动")

		_g.AddText("xm y+12 w320", "自动切换")
		G_PrefAutoSw := _g.AddCheckBox("xm+14 y+6", "新对话框默认开启自动切换")

		_g.AddText("xm y+12 w320", "搜索与历史")
		_g.AddText("xm+14 y+8", "最多保留历史条数：")
		G_PrefMax := _g.AddEdit("x+6 yp-2 w60 Number")

		_g.AddText("xm y+14 w320", "其他")
		_g.AddButton("xm+14 y+6 w120", "打开配置文件").OnEvent("Click", OnPrefsOpenIni)
		_g.AddButton("x+10 w120", "清空搜索历史").OnEvent("Click", OnPrefsClearHistory)

		_g.AddButton("xm+90 y+20 w80 Default", "确定").OnEvent("Click", OnPrefsOK)
		_g.AddButton("x+10 w80", "取消").OnEvent("Click", OnPrefsCancel)

		G_PrefGui := _g
	}
	; 每次打开都按当前设置刷新控件状态
	G_PrefBar.Value := G_bEnableBar ? 1 : 0
	G_PrefAutoSw.Value := G_bDefaultAutoSwitch ? 1 : 0
	G_PrefMax.Value := G_IniConfig.ReadSetting("history_max", 20)
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
		n := 20
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

; Ini
Class IniConfig {
	_filename := ""
	__New(filename) {
		this._filename := filename
	}
	__Delete() {
	}
	ReadDialogNode(DialogStr) {
		return IniRead(this._filename, "Dialogs", DialogStr, "")
	}
	WriteDialogNode(DialogStr, Enable := 1) {
		IniWrite Format("{}={}", DialogStr, Enable), this._filename, "Dialogs"
	}
	ReadDialogPaths(DialogStr) {
		arrPaths := []
		paths := IniRead(this._filename, DialogStr, "paths", "")
		if (paths) {
			arrPaths := StrSplit(paths, "|")
		}
		Loop {
			needCheck := False
			for idx, p in arrPaths {
				if (p = ""){
					arrPaths.RemoveAt(idx)
					needCheck := True
					Break
				}
			}
		} Until !needCheck
		return arrPaths
	}
	WriteDialogPaths(DialogStr, path) {
		arrPaths := this.ReadDialogPaths(DialogStr)
		for p in arrPaths {
			if (p = path){
				Return
			}
		}
		if (arrPaths.Length > 15) {
			arrPaths.Pop()
		}
		arrPaths.InsertAt(1, path)
		_s := ""
		for p in arrPaths {
			if (_s = "")
				_s := p
			Else
				_s := _s . "|" . p 
		}
		IniWrite _s, this._filename, DialogStr, "paths"
	}

	; —— [Settings] 通用设置 ——
	ReadSetting(Key, Default := "") {
		return IniRead(this._filename, "Settings", Key, Default)
	}
	WriteSetting(Key, Value) {
		IniWrite Value, this._filename, "Settings", Key
	}

	; —— [Search] 搜索关键词历史 ——
	ReadSearchHistory() {
		arr := []
		s := IniRead(this._filename, "Search", "history", "")
		if (s != "")
			arr := StrSplit(s, "|")
		return arr
	}
	WriteSearchHistory(arr) {
		s := ""
		for v in arr
			s := (s = "") ? v : s . "|" . v
		IniWrite s, this._filename, "Search", "history"
	}
	; 新增一条历史(去重、置顶、按 history_max 截断)
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
		n := this.ReadSetting("history_max", 20) + 0
		if (n < 1)
			n := 20
		while (arr.Length > n)
			arr.Pop()
		this.WriteSearchHistory(arr)
	}
	ClearSearchHistory() {
		this.WriteSearchHistory([])
	}

}