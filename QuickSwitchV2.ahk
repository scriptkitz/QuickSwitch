#Requires AutoHotkey >=2.0
#Singleinstance Force

#Include "Sqlite.ahk"


SetWorkingDir(A_ScriptDir)

If (VerCompare(A_OSVersion, "10.0.0") <= 0) {
	MsgBox(A_OSVersion " is not supported.")
	ExitApp
}

G_IniConfig := IniConfig("config.ini")

Hotkey("^W", ShowMenu, "Off")

G_DlgHwnd := 0
G_ToolWnd := CreateToolWnd()
_G_LastDlgWndRect := {x:0, y:0, w:0, h:0}
_G_DialogFinger := ""
_G_FDType := ""
_G_DlgIsNotActive := True
_G_bAutoSwitch := True


Loop {
	WinID := WinWaitActive("ahk_class #32770")
	_G_DlgIsNotActive := False
	_G_FDType := GetFileDialogType(WinID)
	if (_G_FDType) {
		G_DlgHwnd := WinID
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
			
		SetTimer _UpToolWndPos, 10
		WinWaitNotActive(WinID)
		G_IniConfig.WriteDialogPaths(_G_DialogFinger, FolderPath)
		_G_DlgIsNotActive := True
		if (!WinActive(G_ToolWnd.Hwnd)) {
			_G_LastDlgWndRect.x := 0
			_G_LastDlgWndRect.y := 0
			_G_LastDlgWndRect.w := 0
			_G_LastDlgWndRect.h := 0
			G_ToolWnd.Hide()
		}
		; Break
	}
}
ExitApp(0)

_UpToolWndPos() {
	global _G_LastDlgWndRect
	if (_G_DlgIsNotActive)
		return
	bShow := WinActive(G_DlgHwnd)
	If (bShow) {
		WinGetPos &x, &y, &w, &h, WinID
		if (_G_LastDlgWndRect.x != x || 
			_G_LastDlgWndRect.y != y || 
			_G_LastDlgWndRect.w != w || 
			_G_LastDlgWndRect.h != h)
		{
			_G_LastDlgWndRect.x := x
			_G_LastDlgWndRect.y := y
			_G_LastDlgWndRect.w := w
			_G_LastDlgWndRect.h := h

			G_ToolWnd.Move(x+(w/2)-25, y)
			G_ToolWnd.Show("NoActivate")
		}
	} else {
		SetTimer ,0
	}
}

OnBtnSetting(Obj, info) {
	AUTOSWITCH := "自动切换"
	OnMenuPath(_name, _pos, _menu) {
		if (AUTOSWITCH = _name) {
			G_IniConfig.WriteDialogNode(_G_DialogFinger, !_G_bAutoSwitch)
			_menu.ToggleCheck(_name)
		} else {
			G_IniConfig.WriteDialogNode(_G_DialogFinger, False)
			FeedDialog(G_DlgHwnd, _name)
		}
	}
	MouseGetPos &x, &y
	m := Menu()
	m.Add(AUTOSWITCH, OnMenuPath)
	_G_bAutoSwitch ? m.Check(AUTOSWITCH) : m.UnCheck(AUTOSWITCH)
	paths := G_IniConfig.ReadDialogPaths(_G_DialogFinger)

	for idx, p in paths {
		m.Add(p, OnMenuPath)
	}
	ControlGetPos &cx, &cy, &cw, &ch, Obj
	m.Show(cx, cy+ch)
}

CreateToolWnd() {
	_gui := Gui("ToolWindow AlwaysOnTop -Sysmenu -Caption -Border -DPIScale")
	_gui.BackColor := "Silver"
	WinSetTransColor "Silver", _gui
	BtnSet := _gui.Add("Button", , "⚙")
	BtnSet.OnEvent("Click", OnBtnSetting)
	BtnSet.Opt("BackgroundTrans")
	_gui.Show("x-100 y-100")
	_gui.Hide()
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
			FeedDialog_SYSLISTVIEW(_WinID, _NewPath)
	}
}

FindAllChild(className, hWnd) {
	GW_HWNDFIRST :=        0
	GW_HWNDLAST :=         1
	GW_HWNDNEXT :=         2
	GW_HWNDPREV :=         3
	GW_OWNER :=            4
	GW_CHILD :=            5
	_child := DllCall("GetWindow", "Ptr", hWnd, "Uint", GW_CHILD|GW_HWNDFIRST)
	Wnds := []
	if(_child) {
		Wnds.Push(FindAllChild(className, _child)*)
		_cls := ControlGetClassNN(_child)
		if (InStr(_cls, className))
			Wnds.Push(_child)
		Loop {
			_child := DllCall("GetWindow", "Ptr", _child, "Uint", GW_HWNDNEXT)
			if(_child) {
				Wnds.Push(FindAllChild(className, _child)*)
				_cls := ControlGetClassNN(_child)
				if (InStr(_cls, className))
					Wnds.Push(_child)
			}
		} Until !_child
	}
	return Wnds
}
; 设置公共对话框的路径
FeedDialog_GENERAL(_WinID, _NewPath) {
	_UseToolbar := ""
	_EnterToolbar := ""
	_SetPathOK := False
	WinActivate(_WinID)
	; ControlFocus "Edit1", _WinID
	_ctrls := WinGetControls(_WinID)
	; 确认确实是公共对话框
	For c in _ctrls {
		if (InStr(c, "ToolbarWindow32")) {
			_hwnd := ControlGetHwnd(c, _WinID)
			_parent := DllCall("GetParent", "Ptr", _hwnd)
			_cls := ControlGetClassNN(_parent)
			if(InStr(_cls, "Breadcrumb Parent")) {
				_UseToolbar := c
			}
			if(InStr(_cls, "msctls_progress32")) {
				_EnterToolbar := c
			}
		}
	}
	If (_UseToolbar AND _EnterToolbar) {
		_loopcount := 0
		Loop {
			_loopcount++
			SendInput "^l" ; 切换地址栏获得焦点 Ctrl+L
			Sleep(100)
			Try {
				_fc := ControlGetFocus("A")
				_fcc := ControlGetClassNN(_fc)
				
				if(InStr(_fcc, "Edit") AND (_fcc != "Edit1"))
				{
					; EditPaste(_NewPath, _fc)
					ControlSetText _NewPath, _fc
					_txt := ControlGetText(_fc)
					if (_txt = _NewPath) {
						_SetPathOK := True
					}
				}
			}
		} Until _SetPathOK OR (_loopcount > 5)
		If (_SetPathOK) {
			ControlSend("{Enter}", _fc)
			ControlClick _EnterToolbar, _WinID
			; 设置回默认的文件名编辑框焦点
			Sleep 10
			ControlFocus "Edit1", _WinID
		}
	} else {
		MsgBox "This type of dialog can not be handled (yet).`nPlease report it!"
	}
}
; 设置Exploer样式的对话框路径
FeedDialog_SYSLISTVIEW(_WinID, _NewPath) {
	WinActivate(_WinID)
	_edit1 := ControlGetHwnd("Edit1", _WinID)
	_oldEditValue := ControlGetText(_edit1)
	_NewPath := RTrim(_NewPath, "\")
	_NewPath := _NewPath . "\"
	_LoopCount := 0
	_FolderSet := False
	Loop {
		_LoopCount++
		Sleep 10
		; 普通权限进程是无法修改管理员权限进程的对话框的
		; EditPaste _NewPath, _edit1
		ControlSetText _NewPath, _edit1
		_txt := ControlGetText(_edit1)
		if(_txt = _NewPath)
			_FolderSet := True
	} Until _FolderSet OR (_LoopCount > 20)
	if(_FolderSet) {
		ControlFocus _edit1
		ControlSend "{Enter}", _edit1
		ControlFocus _edit1
		ControlSetText(_oldEditValue, _edit1)
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

}