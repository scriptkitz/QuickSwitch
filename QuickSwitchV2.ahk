#Requires AutoHotkey >=2.0
#Singleinstance Force

SetWorkingDir(A_ScriptDir)

If (VerCompare(A_OSVersion, "10.0.0") <= 0) {
	MsgBox(A_OSVersion " is not supported.")
	ExitApp
}


Hotkey("^W", ShowMenu, "Off")

Loop {
	WinWaitActive("ahk_class #32770")
	WinID := WinExist("A")
	FDType := GetFileDialogType(WinID)
	if (FDType) {
		windows_title := WinGetTitle(WinID)
		windows_exe := WinGetProcessName(WinID)
		FingerPrint := windows_exe . "___" . windows_title

		FolderPath := GetZWindowPath(WinID)
		if (ValidFolder(FolderPath)) {
			; MsgBox(FDType)
			switch FDType {
				case "GENERAL":
					FeedDialog_GENERAL(WinID, FolderPath)
				case "SYSLISTVIEW":
					FeedDialog_SYSLISTVIEW(WinID, FolderPath)
			}
		}
		Hotkey("^W", "On")
	}

	Sleep(100)
	WinWaitNotActive()
	Hotkey("^W", "Off")
	; Break
}
ExitApp(0)

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
	_zDelta := 2 
	; 一般情况弹出对话框的下一层是主窗口，主窗口的下一个窗口就是我们要找的窗口
	; 所以窗口顺序就是当前对话框+1为主窗口，再+1就是下一个窗口，所以这里默认是2
	WinIDs := WinGetList()
	found_idx := -1
	ZFolder := ""
	Loop WinIDs.Length {
		if (WinIDs[A_Index] = WinID) {
			found_idx := A_Index
		}
	}
	nextWinID := WinIDs[found_idx + _zDelta]
	found_class := WinGetClass(nextWinID)
	Switch found_class, True {
		Case "TTOTAL_CMD": ;	Total Commander
		Case "ThunderRT6FormDC": ;	XYPlorer
		Case "CabinetWClass": ;	File Explorer
			For (ComWin in ComObject("Shell.Application").Windows)
			{
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
	_ctrls := WinGetControls(WinID)
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
			_hwnd := ControlGetHwnd(c)
			_parent := DllCall("GetParent", "Ptr", _hwnd)
			_cls := WinGetClass(_parent)
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
			Sleep(50)
			_fc := ControlGetFocus("A")
			_fcc := WinGetClass(_fc)
			if(InStr(_fcc, "Edit") AND (_fcc != "Edit1"))
			{
				; EditPaste(_NewPath, _fc)
				ControlSetText _NewPath, _fc
				_txt := ControlGetText(_fc)
				if (_txt = _NewPath) {
					_SetPathOK := True
				}
			}
		} Until _SetPathOK OR (_loopcount > 5)
		If (_SetPathOK) {
			ControlClick _EnterToolbar
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