#Requires AutoHotkey v2.0
#SingleInstance Force

; --- CONFIGURATION ---
siteName := "e-courier" 
myExeName := "filter_and_email_reportV2.6.exe"
; ---------------------

SetTimer(WatchForDownload, 500)

WatchForDownload() {
    downloadTitle := "Internet Explorer Download - Security Warning"
    
    if WinExist(downloadTitle) {
        if WinExist(siteName) {
            
            ; 1. Activate the security warning and save
            WinActivate(downloadTitle)
            Sleep(200)
            Send("s") 
            
            ; 2. Wait for the security dialog to actually close
            WinWaitClose(downloadTitle, , 3)
            
            ; 3. Give IE mode a second to finish the hand-off
            Sleep(1000)
            
            ; 4. Close the "about:blank" window
            SetTitleMatchMode(2)
            if WinExist("about:blank") {
                WinClose("about:blank")
            }
            
            ; 5. Close the URL window
            if WinExist("promeddel.e-courier.com") {
                WinClose("promeddel.e-courier.com")
            }

            ; 6. JUMP TO YOUR REPORT EXE
            ; Now outside the previous IF block so it always runs!
            if WinExist("ahk_exe " . myExeName) {
                WinRestore("ahk_exe " . myExeName) ; Restores if minimized
                WinActivate("ahk_exe " . myExeName) ; Brings to front
            }
        }
    }
}

; --- EMERGENCY TERMINATION ---
; Press Ctrl + Shift + Q to instantly kill this script
^+q::ExitApp