#Requires AutoHotkey v2.0
#SingleInstance Force
#Include json.ahk
Persistent


global DEBUG := true
; 設定檔
global google_script_url := IniRead("config.ini", "Settings", "google_script_url")
global SET_INTERVAL := IniRead("config.ini", "Settings", "interval_ms")

TIMER_ID := SetTimer(SyncClipboard, SET_INTERVAL)

; 建立系統匣選單
Tray := A_TrayMenu
Tray.Delete()
Tray.Add("立即同步", SyncClipboard)
Tray.Add("清空剪貼簿", ClearRemoteClipboard)
Tray.Add("退出", (*) => ExitApp())
Tray.Default := "立即同步"
Tray.ClickCount := 1
Tray.Icon := A_WinDir "\System32\shell32.dll", 44

;最後同步的遠端資料
global LastSynced := ""

SyncClipboard(*) {
    url := google_script_url
	global LastSynced
    try {
        clipboardText := HttpGet(url)
        if clipboardText != "" {
			if clipboardText != LastSynced {
                A_Clipboard := clipboardText
                LastSynced := clipboardText
                TrayTip("同步成功", clipboardText, 1)
                Log("同步成功：" clipboardText)
            } else {
                Log("無變更，不更新剪貼簿")
            }
        }
    } catch {
		local empty_response_msg := "從伺服器獲取的剪貼簿內容為空，或者回應格式不正確。"
        MsgBox(empty_response_msg "`n將不繼續嘗試同步。", "同步提示", "iconi")
        Log("同步提示：" empty_response_msg " (空回應或無效資料)")
        StopSync()
    }
}

StopSync() {
    SetTimer SyncClipboard, 0
    Log("同步已被停止")
}

Log(text) {
	global debug
	if (debug) {
		timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
		FileAppend timestamp " - " text "`n", A_ScriptDir "\clipboard_sync.log", "UTF-8"
	}
}

HttpGet(url) {
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", url, false)
    whr.Send()
    json := whr.ResponseText
	Log("伺服器回應內容：" json)
	obj := jxon_load(&json)
	return obj["clipboard"]
}

ClearRemoteClipboard(*) {
    try {
        result := HttpPostClearCommand(google_script_url)
        if InStr(result, "CLEARED") {
            TrayTip("已清空遠端剪貼簿", "", 1)
            Log("使用者清空 Google Sheet")
        } else {
            TrayTip("清空失敗", "未收到正確回應", 3)
			MsgBox(result, "清空失敗", "iconi")
            Log("清空失敗，伺服器回傳：" result)
        }
    } catch {
        local msg := "未知錯誤"
        MsgBox(msg, "清空失敗", "iconi")
        Log("清空失敗：" msg)
    }
}

HttpPostClearCommand(url) {
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("POST", url, false)
    whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    whr.Send("command=CLEAR_CLIPBOARD")
    return whr.ResponseText
}


EncodeURIComponent(str) {
    static enc := ComObject("ScriptControl")
    if !enc.Language
        enc.Language := "JScript"
    return enc.Eval("encodeURIComponent('" str "')")
}
