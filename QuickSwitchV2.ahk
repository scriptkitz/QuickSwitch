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

WM_CHANGEUISTATE := 0x0127
WM_UPDATEUISTATE := 0x0128

G_DlgHwnd := 0
G_LastSetDlg := 0 ; 最后一次自动切换路径的对话框
G_Menu := Menu()
G_ToolWnd := CreateToolWnd()
_G_DialogFinger := ""
_G_FDType := ""
_G_bAutoSwitch := True

Loop {
	WinID := WinWaitActive("ahk_class #32770")
	_G_FDType := GetFileDialogType(WinID)
	if (_G_FDType) {
		G_DlgHwnd := WinID
		UpToolWndParentAndPos(WinID)
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

		WinWaitNotActive(WinID)
		G_IniConfig.WriteDialogPaths(_G_DialogFinger, FolderPath)
		_HideToolWnd(0)
		G_LastSetDlg := 0
	}
}
ExitApp(0)

_HideToolWnd(ParentWnd){
	_style := WinGetStyle(G_ToolWnd.Hwnd)
	_style := _style | WS_POPUP
	_style := _style & ~WS_VISIBLE
	WinSetStyle _style, G_ToolWnd.Hwnd
	preWnd := DllCall("SetParent", "Ptr", G_ToolWnd.Hwnd, "Ptr", ParentWnd)
	flag := SWP_HIDEWINDOW
	DllCall("SetWindowPos", "Ptr", G_ToolWnd.Hwnd, "Ptr", 0, "Int", 0,"Int", 0, "Int", 50, "Int", 50, "UInt", flag, "Int")
}

MAKELONG(LOWORD,HIWORD,Hex:=0){
    BITS:=0x10,WORD:=0xFFFF
    return (!Hex)?((HIWORD<<BITS)|(LOWORD&WORD)):Format("{1:#x}",((HIWORD<<BITS)|(LOWORD&WORD)))
}

_ShowToolWnd(ParentWnd){
	_style := 0
	_style := _style & ~WS_POPUP
	_style := _style | WS_CHILD
	_style := _style | WS_VISIBLE
	WinSetStyle _style, G_ToolWnd.Hwnd
	preWnd := DllCall("SetParent", "Ptr", G_ToolWnd.Hwnd, "Ptr", ParentWnd)
}

UpToolWndParentAndPos(_WinID) {
	WinGetPos &x, &y, &w, &h, _WinID
	WinGetPos &cx, &cy, &cw, &ch, G_ToolWnd.Hwnd
	_parent := DllCall("GetParent", "Ptr", G_ToolWnd.Hwnd)
	if (_parent != _WinID) {
		_ShowToolWnd(_WinID)
	}

	flag := SWP_SHOWWINDOW|SWP_DRAWFRAME
	; 这里的大小如果是50，50，则导致第一次显示正常，后面就显示不出来了，但可以点击到，不知道是否BUG，只好用100，100了
	DllCall("SetWindowPos", "Ptr", G_ToolWnd.Hwnd, "Ptr", 0, "Int", 0,"Int", 0, "Int", 100, "Int", 100, "UInt", flag, "Int")
}

OnBtnSetting(Obj, info) {
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

	_parent := DllCall("GetParent", "Ptr", G_Menu.Handle)
	if (_parent != G_ToolWnd.hwnd) {
		_style := 0
		_style := _style & ~WS_POPUP
		_style := _style | WS_CHILD
		_style := _style | WS_VISIBLE
		WinSetStyle _style, G_ToolWnd.hwnd
		DllCall("SetParent", "Ptr", G_Menu.Handle, "Ptr", G_ToolWnd.hwnd)
	}
	G_Menu.Show(cx, cy+ch)
}

CreateToolWnd() {
	_gui := Gui("")
	_gui.MarginX := 0
	_gui.MarginY := 0
	_gui.BackColor := "Silver"
	_gui.SetFont("s8")
	WinSetTransColor "Silver", _gui
	Btn := _gui.AddButton("-Default -Border w15 h15", "⚙")
	Btn.OnEvent("Click", OnBtnSetting)
	
	return _gui
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
	_GetNextWinID1(WinID) {
		WinIDs := WinGetList()
		_zDelta := 2 
		; 一般情况弹出对话框的下一层是主窗口，主窗口的下一个窗口就是我们要找的窗口
		; 所以窗口顺序就是当前对话框+1为主窗口，再+1就是下一个窗口，所以这里默认是2
		found_idx := -1
		Loop WinIDs.Length {
			if (WinIDs[A_Index] = WinID) {
				found_idx := A_Index
			}
		}
		return WinIDs[found_idx + _zDelta]
	}
	_GetNextWinID2(WinID) {
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

	nextWinID := _GetNextWinID2(WinID)
	found_class := WinGetClass(nextWinID)
	Switch found_class, True {
		Case "TTOTAL_CMD": ;	Total Commander
		Case "ThunderRT6FormDC": ;	XYPlorer
		Case "CabinetWClass": ;	File Explorer
			For (ComWin in ComObject("Shell.Application").Windows)
			{
				_checkID := 0
				Try {
					_checkID := ComWin.hwnd
				} catch Error as err {
				}
				if (nextWinID = _checkID) {
					ZFolder := ComWin.Document.Folder.Self.Path
					Break
				}
			}
		Case "TablacusExplorer": ;	
			For (ComWin in ComObject("Shell.Application").Windows)
			{
				SplitPath ComWin.FullName, &ProcName
				if ((0 = StrCompare(ProcName, "TE32.exe", 0)) or (0 = StrCompare(ProcName, "TE64.exe", 0)))
				{
					if (ComWin.Document.parentWindow)
					{
						ComWin.Document.parentWindow.execScript("document.tophwnd = GetTopWindow(WebBrowser.hwnd)")
						if ( nextWinID = ComWin.Document.tophwnd )
						{
							ZFolder := ComWin.Document.F.addressbar.value
							Break
						}
					}
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
		MsgBox "This type of dialog can not be handled (yet).`nPlease report it!"
		Return
	}
	_edit2 := FindChildByXPath("ComboBoxEx32/ComboBox/Edit", _WinID)
	_NewPath := RTrim(_NewPath, "\")
	_NewPath := _NewPath . "\"
	ControlFocus _edit
	try {
		_oldTxt := ControlGetText(_edit,_WinID)
	} catch {
		Return
	}
	ControlSetText _NewPath, _edit, _WinID
	ControlSend("{Enter}", _edit, _WinID)
	; 设置回默认的文件名编辑框焦点
	ControlSetText _oldTxt, _edit, _WinID
}
; 设置公共对话框的路径
FeedDialog_Test(_WinID, _NewPath) {
	_AddressToolbar := ""
	WinActivate(_WinID)
	Sleep(2000)
	if !WinExist(_WinID) {
		return
	}
	_ctrls := WinGetControls(_WinID)
	; 获取地址栏控件
	For c in _ctrls {
		if (InStr(c, "ToolbarWindow32")) {
			_hwnd := ControlGetHwnd(c, _WinID)
			_parent := DllCall("GetParent", "Ptr", _hwnd)
			_cls := ControlGetClassNN(_parent, _WinID)
			if(InStr(_cls, "Breadcrumb Parent")) {
				_AddressToolbar := c
			}
		}
	}
	If (!_AddressToolbar) {
		MsgBox "This type of dialog can not be handled (yet).`nPlease report it!"
		Return
	}
	; 循环等待Edit2获得焦点
	Loop {
		;ControlClick _AddressToolbar, _WinID,,,,"X15 Y15" ; 点击地址栏，等待隐藏的控件输入框可见并获得焦点
		; 不使用Ctrl+l获得地址栏输入焦点了，直接通过控件点击
		SendInput "{Ctrl down}{l down}" ; 切换地址栏获得焦点 Ctrl+L
		SendInput "{Ctrl up}"
		_fc := ControlGetFocus(_WinID)
		if (!_fc) {
			return
		}
		_fcc := ControlGetClassNN(_fc, _WinID)
		if _fcc = "Edit2" {
			; if(1) {
			; 	ControlFocus "Edit1", _WinID
			; 	return
			; }
			ControlSetText _NewPath, _fc, _WinID
			; Sleep(500)
			if (ControlGetText(_fc,_WinID) = _NewPath) {
				
				ControlFocus _fc, _WinID
				ControlSend("{Enter}", _fcc, _WinID)
				; 设置回默认的文件名编辑框焦点
				ControlFocus "Edit1", _WinID
				Break
			}
		}
	} Until !WinActive(_WinID)
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