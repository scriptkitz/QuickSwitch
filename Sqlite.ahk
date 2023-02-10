CreateTableSQL := "BEGIN TRANSACTION;" .
"CREATE TABLE IF NOT EXISTS config (ID INTEGER PRIMARY KEY AUTOINCREMENT, 'Key' TEXT UNIQUE NOT NULL, Value TEXT);" .
"CREATE TABLE IF NOT EXISTS dialog (ID INTEGER PRIMARY KEY AUTOINCREMENT, dialog TEXT UNIQUE NOT NULL);" .
"CREATE TABLE IF NOT EXISTS dialog_paths (ID INTEGER PRIMARY KEY AUTOINCREMENT, dialog_id INTEGER REFERENCES dialog (ID), path TEXT);" .
"COMMIT TRANSACTION;"

Class SQliteDB {
    Static _MinVersion := "3.6"
    DLL := 0
    Version := ""
    _SQLiteDLL := A_ScriptDir . "\SQLite3.dll"
    _DBPath := ""
    _DBHandle := 0
    ErrorMsg := ""
    ErrorCode := 0
    SQL := ""
    Changes := 0

    __New(sqlite_dll) {
        This.Base._SQLiteDLL := sqlite_dll
        If !(This.DLL := DllCall("LoadLibrary", "Str", This.Base._SQLiteDLL, "UPtr")) {
            MsgBox "DLL does not exist!","SQLiteDB Error", "Iconx"
            ExitApp
        }
        This.Base.Version := StrGet(DllCall("SQlite3.dll\sqlite3_libversion", "Cdecl UPtr"), "UTF-8")
        If(VerCompare(This.Base.Version, SQliteDB._MinVersion) < 0)
        {
            MsgBox This.Base.Version .  " of SQLite3.dll is not supported!`n"
                                      . "You can download the current version from www.sqlite.org!", "ERROR", "Iconx"
            ExitApp
        }
    }
    __Delete() {
        If (This.DLL != 0) {
            DllCall("FreeLibrary", "Ptr", This.DLL)
            This.DLL := 0
        }
    }
    _StrToUTF8Buffer(Str) {
        _buf := Buffer(StrPut(Str, "UTF-8"))
        StrPut(Str, _buf, "UTF-8")
        return _buf
    }
    _UTF8BufferToStr(UTF8) {
        Return StrGet(UTF8, "UTF-8")
    }
    _EscapeStr(Str, Quote := True) {
        _buf := this._StrToUTF8Buffer(Str)
        _fbuf := Buffer(3)
        StrPut(Quote ? "%Q" : "%q", _fbuf, "UTF-8")
        Ptr := DllCall("SQLite3.dll\sqlite3_mprintf", "Ptr", _fbuf, "Ptr", _buf, "Cdecl UPtr")
        rStr := this._UTF8BufferToStr(Ptr)
        DllCall("SQLite3.dll\sqlite3_free", "Ptr", Ptr, "Cdecl")
        return rStr
    }
    OpenDB(db_path) {
        Static SQLITE_OPEN_READONLY  := 0x01 ; Database opened as read-only
        Static SQLITE_OPEN_READWRITE := 0x02 ; Database opened as read-write
        Static SQLITE_OPEN_CREATE    := 0x04 ; Database will be created if not exists
        ; Static MEMDB := ":memory:"
        This.ErrorMsg := ""
        This.ErrorCode := 0
        HDB := 0
        if ((db_path = this._DBPath) && this._DBHandle)
            return True
        if (this._DBHandle) {
            this.ErrorMsg := "You must first close DB."
            return False
        }
        this._DBPath := db_path
        Flags := SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE
        UTF8 := ""
        buff := this._StrToUTF8Buffer(this._DBPath)
        RC := DllCall("SQlite3.dll\sqlite3_open_v2", "Ptr", buff, "UPtrP", &HDB, "Int", Flags, "Ptr", 0, "Cdecl Int")
        if(RC) {
            this._DBPath := ""
            this.ErrorMsg := ""
            this.ErrorCode := 0
            return False
        }
        this._DBHandle := HDB
        this.Exec(CreateTableSQL)
        return True
    }
    CloseDB() {
        this.ErrorMsg := ""
        this.ErrorCode := 0
        this.SQL := ""
        if !(this._DBHandle)
            return True
        DllCall("SQlite3.dll\sqlite3_close", "Ptr", This._DBHandle, "Cdecl Int")
    }
    Exec(SQL, Callback := "") {
        if !(this._DBHandle) {
            this.ErrorMsg := "Invalid database handle!"
            return False
        }
        CBPtr := 0
        Err := 0
        if (Callback != "" && Callback.MinParams = 4)
            CBPtr := CallbackCreate(Callback, "F C", 4)
        _ustr := this._StrToUTF8Buffer(SQL)
        RC := DllCall("SQlite3.dll\sqlite3_exec", "Ptr", this._DBHandle, "Ptr", _ustr, "Int", CBPtr, "Ptr", ObjPtr(This), "UPtrP", &Err, "Cdecl Int")
        if (CBPtr)
            CallbackFree(CBPtr)
        if (RC) {
            this.ErrorMsg := StrGet(Err, "UTF-8")
            this.ErrorCode := RC
            DllCall("SQLite3.dll\sqlite3_free", "Ptr", Err, "Cdecl")
            return False
        }
        This.Changes := DllCall("SQLite3.dll\sqlite3_changes","Ptr", this._DBHandle, "Cdecl Int")
        return True
    }

    AddDialogNode(value) {
        _cb(ud, colCount, colValues, colNames) {
            if (colCount > 0)
                _id := StrGet(NumGet(colValues, "UInt"), "UTF-8")
        }
        ev := this._EscapeStr(value)
        SQL := Format("INSERT INTO dialog (dialog) VALUES ({});", ev)
        if (this.Exec(SQL, _cb)) {
            return this.GetDialogNodeID(value)
        }
        return ""
    }
    GetDialogNodeID(value) {
        _id := ""
        _cb(ud, colCount, colValues, colNames) {
            if (colCount > 0)
                _id := StrGet(NumGet(colValues, "UInt"), "UTF-8")
        }
        _sql := Format("SELECT ID FROM dialog WHERE dialog='{}';", value)
        this.Exec(_sql, _cb)
        return _id
    }
}
DialogID := 0
SqlCB(userdata, ColumnCount, ColumnValues, ColumnNames) {
    SqliteObj := ObjFromPtrAddRef(userdata)
}
OnInsertDialog(userdata, ColumnCount, ColumnValues, ColumnNames) {
    global DialogID
    SqliteObj := ObjFromPtrAddRef(userdata)
    Loop ColumnCount {
        idx := A_Index-1
        DialogID := StrGet(NumGet(ColumnValues + idx, "UInt"), "UTF-8")
        cn := StrGet(NumGet(ColumnNames + idx, "UInt"), "UTF-8")
    }
}






; a := SQliteDB("Sqlite/Sqlite3.dll")
; a.OpenDB("./config.db")





; a.Exec(CreateTableSQL, SqlCB)
; b := a.GetDialogNodeID("E:\OneDrive\GreenSoft\AutoHotKey1.x\Scripts\QuickSwitch")
; ; a.AddDialogNode("E:\OneDrive\GreenSoft\AutoHotKey1.x\Scripts\QuickSwitch", SqlCB)
; a.CloseDB()

; ExitApp 0