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

G_DlgHwnd := 0
G_LastSetDlg := 0 ; 最后一次自动切换路径的对话框
G_Menu := Menu()
G_ComboHwnd := 0
G_BarBtnHwnd := 0
G_BarGui := CreateBarWnd()
G_BarHwnd := G_BarGui.Hwnd
G_SearchText := ""
; 注册全局窗口事件钩子：监听文件对话框自身的位置/尺寸变化，同步搜索框位置。
; 用 WINEVENT_OUTOFCONTEXT(默认 flags=0) 以便跨进程捕获其它程序的对话框事件。
G_WinEventProc := CallbackCreate(WinEventProc, "Fast", 7)
G_WinEventHook := DllCall("SetWinEventHook", "UInt", EVENT_OBJECT_LOCATIONCHANGE, "UInt", EVENT_OBJECT_LOCATIONCHANGE, "Ptr", 0, "Ptr", G_WinEventProc, "UInt", 0, "UInt", 0, "UInt", 0, "Ptr")
_G_DialogFinger := ""
_G_FDType := ""
_G_bAutoSwitch := True

Loop {
	WinID := WinWaitActive("ahk_class #32770")
	_G_FDType := GetFileDialogType(WinID)
	if (_G_FDType) {
		G_DlgHwnd := WinID
		ShowBar(WinID)
		; 激活的窗口在失去激活前只能自动设置一次路径。
		if (G_LastSetDlg = G_DlgHwnd) { 
			Continue
		}
		windows_title := WinGetTitle(WinID)
		windows_exe := WinGetProcessName(WinID)
		_G_DialogFinger := windows_exe . "___" . windows_title
		
		_bAuto := G_IniConfig.ReadDialogNode(_G_DialogFinger)
		if (_bAuto || _bAuto == "") {
			_G_bAutoSwitch := True
		} else {
			_G_bAutoSwitch := False
		}
		if(_bAuto == "") {
			G_IniConfig.WriteDialogNode(_G_DialogFinger)
		}
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
	OnMenuPath(_name, _pos, _menu) {
		if (AUTOSWITCH = _name) {
			_menu.ToggleCheck(_name)
			G_IniConfig.WriteDialogNode(_G_DialogFinger, !_G_bAutoSwitch)
		} else {
			FeedDialog(G_DlgHwnd, _name)
		}
	}
	MouseGetPos &x, &y
	G_Menu.Delete()
	G_Menu.Add(AUTOSWITCH, OnMenuPath)
	_G_bAutoSwitch ? G_Menu.Check(AUTOSWITCH) : G_Menu.UnCheck(AUTOSWITCH)
	paths := G_IniConfig.ReadDialogPaths(_G_DialogFinger)

	for idx, p in paths {
		G_Menu.Add(p, OnMenuPath)
	}
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
	global G_ComboHwnd, G_BarBtnHwnd
	_gui := Gui("+ToolWindow -Caption +AlwaysOnTop")
	_gui.MarginX := 4
	_gui.MarginY := 3
	_gui.BackColor := "F2F2F2"
	; ComboBox：以后用 Add()/Delete() 放输入历史，可下拉选择（h200 为下拉列表高度）
	_combo := _gui.AddComboBox("w280 h200", [])
	G_ComboHwnd := _combo.Hwnd
	_combo.OnEvent("Change", OnSearchInput)
	; ⚙ 按钮（位置/尺寸由 RepositionBar 统一布局；未来承载菜单/配置）
	_btn := _gui.AddButton("x300 y3 w32 h24", "⚙")
	G_BarBtnHwnd := _btn.Hwnd
	_btn.OnEvent("Click", OnBtnSetting)
	; 隐藏默认按钮：捕获 Enter（ComboBox 没有 KeyDown 事件）
	_hbtn := _gui.AddButton("Default Hidden w1 h1", "")
	_hbtn.OnEvent("Click", OnSearchSubmit)
	return _gui
}

; 显示底部工具条。
ShowBar(_WinID) {
	global G_BarGui, G_BarHwnd
	if (!G_BarGui)
		return
	G_BarHwnd := G_BarGui.Hwnd
	RepositionBar(_WinID)
}

HideBar() {
	global G_BarHwnd
	if (G_BarHwnd && DllCall("IsWindow", "Ptr", G_BarHwnd))
		DllCall("SetWindowPos", "Ptr", G_BarHwnd, "Ptr", 0, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", SWP_HIDEWINDOW | SWP_NOACTIVATE, "Int")
	G_BarHwnd := 0
}

; 把工具条贴到对话框下边缘外侧（屏幕坐标），并排好内部 [ ComboBox | ⚙ ]。
; 可被 Fast 回调直接调用：只用 DllCall/Buffer/NumGet。
RepositionBar(_WinID) {
	global G_BarHwnd, G_ComboHwnd, G_BarBtnHwnd, BAR_H, BAR_BTN_W, BAR_GAP
	if (!G_BarHwnd)
		return
	if (!DllCall("IsWindow", "Ptr", _WinID))
		return
	_rect := Buffer(16)
	if (!DllCall("GetWindowRect", "Ptr", _WinID, "Ptr", _rect))
		return
	_l := NumGet(_rect, 0, "Int")
	_r := NumGet(_rect, 8, "Int")
	_b := NumGet(_rect, 12, "Int")
	_w := _r - _l
	if (_w < 200)
		_w := 200
	; 工具条窗：贴在对话框下边缘外侧，宽度与对话框一致，用 SWP_NOZORDER 保持 AlwaysOnTop 层位
	DllCall("SetWindowPos", "Ptr", G_BarHwnd, "Ptr", 0, "Int", _l, "Int", _b, "Int", _w, "Int", BAR_H, "UInt", SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW | SWP_NOSENDCHANGING, "Int")
	; ComboBox 可见行高(CB_GETITEMHEIGHT -1)，仅用于让 ComboBox 垂直居中
	_row_h := 22
	if (G_ComboHwnd && DllCall("IsWindow", "Ptr", G_ComboHwnd)) {
		_eh := DllCall("SendMessage", "Ptr", G_ComboHwnd, "UInt", 0x0154, "Ptr", -1, "Ptr", 0, "Ptr")
		if (_eh >= 16 && _eh <= BAR_H - 4)
			_row_h := _eh
	}
	; ⚙ 按钮固定高度(不被压扁)；两者各自垂直居中 => 中心轴一致即视觉对齐
	_btn_h := BAR_H - 6
	_cy := (BAR_H - _row_h) // 2
	_by := (BAR_H - _btn_h) // 2
	; ComboBox：占满左侧（只改宽度与纵向位置，高度沿用创建值，避免破坏下拉列表高度）
	if (G_ComboHwnd && DllCall("IsWindow", "Ptr", G_ComboHwnd)) {
		_crect := Buffer(16)
		DllCall("GetWindowRect", "Ptr", G_ComboHwnd, "Ptr", _crect)
		_chh := NumGet(_crect, 12, "Int") - NumGet(_crect, 4, "Int")
		_cw := _w - 8 - BAR_BTN_W - BAR_GAP
		DllCall("SetWindowPos", "Ptr", G_ComboHwnd, "Ptr", 0, "Int", 4, "Int", _cy, "Int", _cw, "Int", _chh, "UInt", SWP_NOZORDER | SWP_NOACTIVATE, "Int")
	}
	; ⚙ 按钮：右侧，固定高度，与 ComboBox 中心轴对齐
	if (G_BarBtnHwnd && DllCall("IsWindow", "Ptr", G_BarBtnHwnd))
		DllCall("SetWindowPos", "Ptr", G_BarBtnHwnd, "Ptr", 0, "Int", _w - 4 - BAR_BTN_W, "Int", _by, "Int", BAR_BTN_W, "Int", _btn_h, "UInt", SWP_NOZORDER | SWP_NOACTIVATE, "Int")
}

; 等待"对话框会话"结束：对话框被销毁，或对话框与工具条都不再是前台窗口。
; 关键：把焦点落在工具条上视为会话仍在进行，否则点一下工具条就会把它自己隐藏掉。
WaitDialogSessionEnd(WinID) {
	global G_BarHwnd
	Loop {
		if (WinActive("ahk_id " WinID)) {
			Sleep 30
			continue
		}
		; 焦点落在工具条上时视为会话仍在进行，
		; 但需同时确认对话框还在，否则会因工具条仍在前台而永远等不到结束
		if (G_BarHwnd && WinActive("ahk_id " G_BarHwnd) && DllCall("IsWindow", "Ptr", WinID)) {
			Sleep 30
			continue
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

OnSearchInput(GuiCtrlObj, Info) {
	global G_SearchText
	G_SearchText := GuiCtrlObj.Text
	; 未来在此根据 G_SearchText 查询文件
}

; 隐藏默认按钮的 Click：搜索框中按 Enter 时触发（未来接文件查询）
OnSearchSubmit(GuiCtrlObj, Info) {
	global G_SearchText
	ToolTip("待实现查询: " G_SearchText)
	SetTimer(() => ToolTip(), -1500)
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

}