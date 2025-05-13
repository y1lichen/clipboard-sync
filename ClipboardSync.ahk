#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

DEBUG := false
SET_INTERVAL := 5000  ; 每 5 秒同步一次
TIMER_ID := SetTimer(SyncClipboard, SET_INTERVAL)

; 建立系統匣選單
Tray := A_TrayMenu
Tray.Delete()
Tray.Add("立即同步", SyncClipboard)
Tray.Add("退出", (*) => ExitApp())
Tray.Default := "立即同步"
Tray.ClickCount := 1
Tray.Icon := A_WinDir "\System32\shell32.dll", 44
google_script_url := "https://script.google.com/macros/s/AKfycbzyqOCM534vFDlokCxn_5Lt3Ff1kIXGDEwfi42VWRGRZ7M5Dopt-nhzXH_PDwgleYtw9w/exec"

SyncClipboard(*) {
    url := google_script_url
    try {
        clipboardText := HttpGet(url)
        if clipboardText != "" {
            A_Clipboard := clipboardText
            TrayTip("同步成功", clipboardText, 5)
            Log("同步成功：" clipboardText)
        } else {
            TrayTip("同步失敗", "空回應，停止同步", 5)
            Log("同步失敗：空回應")
            StopSync()
        }
    } catch {
		local empty_response_msg := "從伺服器獲取的剪貼簿內容為空，或者回應格式不正確。"
        MsgBox(empty_response_msg "`n將不繼續嘗試同步。", "同步提示", "iconi")
        Log("同步提示：" empty_response_msg " (空回應或無效資料)")
        StopSync()
    }
}

StopSync() {
    SetTimer(SyncClipboard, 0)
    Log("同步已被停止")
}

Log(text) {
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

    ; 從 JSON 取出 clipboard 欄位
    if RegExMatch(json, '"clipboard"\s*:\s*"(.+?)"', &match)
        return StrReplace(match[1], "\\n", "`n")
    return ""
}
