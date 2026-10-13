; ============================================================================
; Everything3.ahk —— voidtools "Everything 3" (SDK3) 的最小 AHK 封装
; ----------------------------------------------------------------------------
; 只做一件事：拿关键词去 Everything 里搜路径。不依赖本工程任何变量，可单独复用。
;
; 依赖：Everything3.dll（与脚本同目录，或已在 PATH / 已 LoadLibrary），
;       且 Everything 客户端正在运行（SDK3 本质是命名管道客户端）。
;
; 用法：
;     res := Ev3_Search("报告 2024")     ; 默认最多 60 条
;     !IsObject(res) → 不可用（Everything3.dll 缺失 或 Everything 未运行），此时为 0
;     res.Length = 0 → 搜到了但没有结果
;     否则           → 命中的完整路径数组，如 "D:\docs\报告 2024.docx"
;
; 说明：连接会缓存复用（管道可长开）；若 Everything 重启导致连接失效，
;       内部会自动丢弃缓存，下次调用重连。
; ============================================================================

EV3_DLL    := "Everything3"   ; DllCall 前缀：Everything3\Everything3_XxxW
EV3_MAX    := 50              ; 默认最多返回多少条（每条取值约 2ms，条数直接决定耗时）
EV3_CLIENT := 0               ; 缓存的客户端句柄（0 = 未连接）

; 连接 Everything：先默认实例，失败再试 1.5a（alpha 实例）；失败返回 0
Ev3_Connect() {
	global EV3_DLL, EV3_CLIENT
	if (EV3_CLIENT)
		return EV3_CLIENT
	c := DllCall(EV3_DLL "\Everything3_ConnectW", "Ptr", 0, "Ptr")
	if (!c)
		c := DllCall(EV3_DLL "\Everything3_ConnectW", "WStr", "1.5a", "Ptr")
	EV3_CLIENT := c
	return c
}

; 丢弃缓存的连接
Ev3_Reset() {
	global EV3_DLL, EV3_CLIENT
	if (EV3_CLIENT) {
		try DllCall(EV3_DLL "\Everything3_DestroyClient", "Ptr", EV3_CLIENT, "Int")
		EV3_CLIENT := 0
	}
}

; 搜索：命中返回全路径数组；0 = 不可用；[] = 无结果
Ev3_Search(query, maxResult := 0) {
	global EV3_DLL, EV3_CLIENT, EV3_MAX
	if (maxResult <= 0)
		maxResult := EV3_MAX
	try
		client := Ev3_Connect()
	catch
		return 0                                   ; DLL 不存在
	if (!client)
		return 0                                   ; Everything 未运行
	try {
		ss := DllCall(EV3_DLL "\Everything3_CreateSearchState", "Ptr")
		if (!ss)
			return 0
		DllCall(EV3_DLL "\Everything3_SetSearchTextW", "Ptr", ss, "WStr", query, "Int")
		DllCall(EV3_DLL "\Everything3_SetSearchViewportCount", "Ptr", ss, "UPtr", maxResult, "Int")
		rl := DllCall(EV3_DLL "\Everything3_Search", "Ptr", client, "Ptr", ss, "Ptr")
		out := []
		if (rl) {
			; ★ 只能迭代"视口数量"(=我们请求的条数)，绝不能迭代 GetResultListCount(总匹配数)：
			;   视口之外的条目每次取值仍是一次完整调用却返回空，广查询(如单字母)会因此卡几秒~几十秒。
			n := DllCall(EV3_DLL "\Everything3_GetResultListViewportCount", "Ptr", rl, "UPtr")
			; 复用一块缓冲：单次调用直接取字符串。
			; 比"先问所需长度、再分配、再取"的两段式少一半调用与每行的内存分配。
			cap := 1024                                  ; wchar 数，普通路径远小于此
			buf := Buffer(cap * 2)                       ; wchar = 2 字节
			Loop n {
				i := A_Index - 1
				need := DllCall(EV3_DLL "\Everything3_GetResultFullPathNameW", "Ptr", rl, "UPtr", i, "Ptr", buf, "UPtr", cap, "UPtr")
				if (!need)
					continue
				if (need > cap) {                        ; 超长路径：按需扩一次
					cap := need
					buf := Buffer(cap * 2)
					DllCall(EV3_DLL "\Everything3_GetResultFullPathNameW", "Ptr", rl, "UPtr", i, "Ptr", buf, "UPtr", cap, "UPtr")
				}
				out.Push(StrGet(buf))
			}
			DllCall(EV3_DLL "\Everything3_DestroyResultList", "Ptr", rl, "Int")
		} else if (Ev3_IsConnError()) {
			Ev3_Reset()                            ; 连接已失效 → 下次重连
		}
		DllCall(EV3_DLL "\Everything3_DestroySearchState", "Ptr", ss, "Int")
		return out
	} catch {
		return 0
	}
}

; 最近一次 SDK 调用的错误码（EVERYTHING3_OK / EVERYTHING3_ERROR_*）
Ev3_LastError() {
	global EV3_DLL
	try
		return DllCall(EV3_DLL "\Everything3_GetLastError", "UInt")
	catch
		return -1
}

; 是否为"连接类"错误（管道不存在 / 已断开）
Ev3_IsConnError() {
	err := Ev3_LastError()
	return err = 0xE0000002 || err = 0xE0000003
}
