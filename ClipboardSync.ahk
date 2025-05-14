#Requires AutoHotkey v2.0
#SingleInstance Force
#Include json.ahk
Persistent

global DEBUG := false
; Config
global google_script_url := IniRead("config.ini", "Settings", "google_script_url")
global SET_INTERVAL := IniRead("config.ini", "Settings", "interval_ms")
global max_clipboard_length := IniRead("config.ini", "Settings", "max_clipboard_length")

TIMER_ID := SetTimer(SyncClipboard, SET_INTERVAL)

; System Tray Menu
Tray := A_TrayMenu
Tray.Delete()
Tray.Add("Sync Now", SyncClipboard)
Tray.Add("Clear Remote Clipboard", ClearRemoteClipboard)
Tray.Add("Upload Clipboard", UploadClipboardToGoogleSheet)
Tray.Add("Exit", (*) => ExitApp())
Tray.Default := "Sync Now"
Tray.ClickCount := 1
Tray.Icon := A_WinDir "\System32\shell32.dll", 44

global LastSynced := ""
global LastUploaded := ""

OnClipboardChange(ClipboardChanged)

SyncClipboard(*) {
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
                TrayTip("Sync Successful", clipboardText, 1)
                Log("Sync successful: " clipboardText)
            } else {
                Log("No changes, clipboard not updated")
            }
        }
		if (max_clipboard_length > 0 and rowCount > max_clipboard_length) {
            Log("Row count reached " rowCount ", auto clearing")
            ClearRemoteClipboard()
        }
    } catch {
		local empty_response_msg := "Clipboard content is empty or invalid response format."
        MsgBox(empty_response_msg "`nSync aborted.", "Sync Notice", "iconi")
        Log("Sync notice: " empty_response_msg " (Empty or invalid response)")
        StopSync()
    }
}

StopSync() {
    SetTimer SyncClipboard, 0
    Log("Sync has been stopped")
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
	Log("Server response: " json)
	obj := jxon_load(&json)
	return obj
}

ClipboardChanged(*) {
    global LastUploaded, google_script_url

    text := A_Clipboard

    if text = "" || text = LastUploaded
        return

    try {
        result := HttpPostClipboard(google_script_url, text)
        if InStr(result, "OK") {
            LastUploaded := text
            Log("Auto upload successful: " text)
        } else {
            TrayTip("Upload Failed", result, 3)
            Log("Auto upload failed: " result)
        }
    } catch {
        Log("Upload error: Exception during auto upload")
    }
}

ClearRemoteClipboard(*) {
    try {
        result := HttpPostClearCommand(google_script_url)
        if InStr(result, "CLEARED") {
            TrayTip("Remote clipboard cleared", "", 1)
            Log("User cleared Google Sheet clipboard")
        } else {
            TrayTip("Clear Failed", "Invalid server response", 3)
            Log("Clear failed, server response: " result)
        }
    } catch {
        local msg := "Unknown error"
        MsgBox(msg, "Clear Failed", "iconi")
        Log("Clear failed: " msg)
    }
}

UploadClipboardToGoogleSheet(*) {
    text := A_Clipboard
    if (text != "") {
        try {
            result := HttpPostClipboard(google_script_url, text)
            if InStr(result, "OK") {
                Log("Manual upload successful: " text)
            } else {
                TrayTip("Upload Failed", "Server did not return OK", 3)
                Log("Manual upload failed, server response: " result)
            }
        } catch {
            local msg := "Error during upload"
            MsgBox(msg, "Upload Error", "iconx")
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
