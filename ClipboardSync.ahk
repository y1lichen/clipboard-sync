#Requires AutoHotkey v2.0
#SingleInstance Force
#Include json.ahk
Persistent

global DEBUG := IniRead("config.ini", "Settings", "debug", "false") = "true"

; 設定檔
global google_script_url := IniRead("config.ini", "Settings", "google_script_url")
global SET_INTERVAL := IniRead("config.ini", "Settings", "interval_ms")
global max_clipboard_length := IniRead("config.ini", "Settings", "max_clipboard_length")

TIMER_ID := SetTimer(() => SyncClipboard(false), SET_INTERVAL)

; 建立系統匣選單
Tray := A_TrayMenu
Tray.Delete()
Tray.Add("Sync Now", (*) => SyncClipboard(true))
Tray.Add("Clear Remote Clipboard", (*) => ClearRemoteClipboard(true))
Tray.Add("Upload Clipboard", (*) => UploadClipboardToGoogleSheet(true))
Tray.Add("Exit", (*) => ExitApp())
Tray.Default := "Sync Now"
Tray.ClickCount := 1
Tray.Icon := A_WinDir "\System32\shell32.dll", 44

global upload_pending := false
; delay in ms
global upload_delay := -150


global LastSynced := ""
global LastUploaded := ""

OnClipboardChange(ClipboardChanged)

SyncClipboard(showTip := false) {
    url := google_script_url
    global LastSynced, max_clipboard_length
    try {
        obj := HttpGet(url)

        clipboardText := obj["clipboard"]
        rowCount := obj["rowCount"]

        if clipboardText != "" {
            if clipboardText != LastSynced {
                A_Clipboard := clipboardText
                LastSynced := clipboardText
                if (showTip) {
                    TrayTip("Sync Successful", clipboardText, 1)
                }
                Log("Sync successful: " clipboardText)
            } else {
                Log("No changes, clipboard not updated")
            }
        }
        if (max_clipboard_length > 0 and rowCount > max_clipboard_length) {
            Log("Row count reached " rowCount ", auto clearing")
            ClearRemoteClipboard(false)
        }
    } catch {
        local empty_response_msg := "Clipboard content is empty or invalid response format."
        if (showTip) {
            MsgBox(empty_response_msg "`nSync aborted.", "Sync Notice", "iconi")
        }
        Log("Sync notice: " empty_response_msg " (Empty or invalid response)")
        StopSync()
    }
}

StopSync() {
    SetTimer SyncClipboard, 0
    Log("Sync stopped")
}

Log(text) {
    global DEBUG
    if (DEBUG) {
        timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        FileAppend timestamp " - " text "`n", A_ScriptDir "\clipboard_sync.log", "UTF-8"
    }
}

HttpGet(url) {
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", url, false)
    whr.Send()
    json := whr.ResponseText
    Log("Server response: " json)
    obj := jxon_load(&json)
    return obj
}

ClipboardChanged(*) {
    if !ClipboardHasText()
        return
    global LastUploaded, google_script_url, upload_pending, upload_delay

    text := A_Clipboard

    if text = "" || text = LastUploaded
        return

	; 避免短時間內多次觸發
    if !upload_pending {
        upload_pending := true
        SetTimer () => DoClipboardUpload(text), upload_delay  ; 傳遞 text 延遲處理
    }
}

ClipboardHasText() {
    try {
		return !!StrLen(A_Clipboard)
	}
    catch {
		return false
	}
}

DoClipboardUpload(text) {
    global upload_pending, LastUploaded
    upload_pending := false
    result := HttpPostClipboard(google_script_url, text)
    if InStr(result, "OK") {
        LastUploaded := text
        Log("Auto upload successful: " text)
    } else {
        Log("Auto upload failed: " result)
    }
}

ClearRemoteClipboard(showTip := false) {
    try {
        result := HttpPostClearCommand(google_script_url)
        if InStr(result, "CLEARED") {
            if (showTip) {
                TrayTip("Remote clipboard cleared", "", 1)
            }
            Log("Remote clipboard cleared")
        } else {
            if (showTip) {
                TrayTip("Clear failed", "Invalid server response", 3)
            }
            Log("Clear failed, server response: " result)
        }
    } catch {
        local msg := "Unknown error"
        if (showTip) {
            MsgBox(msg, "Clear Failed", "iconi")
        }
        Log("Clear failed: " msg)
    }
}

UploadClipboardToGoogleSheet(showTip := true) {
    text := A_Clipboard
    if (text != "") {
        try {
            result := HttpPostClipboard(google_script_url, text)
            if InStr(result, "OK") {
                if (showTip) {
                    TrayTip("Upload Successful", "", 1)
                }
                Log("Manual upload successful: " text)
            } else {
                if (showTip) {
                    TrayTip("Upload Failed", "Server did not return OK", 3)
                }
                Log("Manual upload failed, server response: " result)
            }
        } catch {
            local msg := "Error during upload"
            if (showTip) {
                MsgBox(msg, "Upload Error", "iconx")
            }
            Log("Upload error: " msg)
        }
    }
}

HttpPostClipboard(url, text) {
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("POST", url, false)
    whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    postData := "clipboard=" EncodeURIComponent(text)
    whr.Send(postData)
    response := whr.ResponseText
    Log("POST response: " response)
    return response
}

HttpPostClearCommand(url) {
    whr := ComObject("WinHttp.WinHttpRequest.5.1")
    whr.Open("POST", url, false)
    whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    whr.Send("command=CLEAR_CLIPBOARD")
    return whr.ResponseText
}

EncodeURIComponent(Url, Flags := 0x000C3000) {
    Local CC := 4096, Esc := "", Result := ""
    Loop
        VarSetStrCapacity(&Esc, CC), Result := DllCall("Shlwapi.dll\UrlEscapeW", "Str", Url, "Str", &Esc, "UIntP", &CC, "UInt", Flags, "UInt")
    Until Result != 0x80004003 ; E_POINTER
    Return Esc
}
