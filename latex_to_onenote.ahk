; LaTeX to OneNote Converter Launcher
; This script runs latex_to_onenote.py using the configured virtual environment.

#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%

; Paths
EnvGet, UserProfile, USERPROFILE
VenvPython := UserProfile . "\.virtualenvs\latex_to_onenote_venv\Scripts\pythonw.exe"
ScriptFile := A_ScriptDir . "\latex_to_onenote.py"

; Check if environment exists
if !FileExist(VenvPython) {
    MsgBox, 16, LaTeX to OneNote Error, Python virtual environment not found at:`n%VenvPython%`n`nPlease run 'setup_env.ps1' to create the environment first.
    ExitApp
}

; Configuration
AutoPaste := 1

; Tray Menu Setup
Menu, Tray, Tip, LaTeX to OneNote Converter
Menu, Tray, Add, Show Shortcut, ShowReminder
Menu, Tray, Add, Auto-Paste on Success, ToggleAutoPaste
Menu, Tray, Check, Auto-Paste on Success
Menu, Tray, Default, Show Shortcut ; Make clicking the icon trigger the popup
Menu, Tray, Click, 1 ; Trigger on single click

; Startup Reminder
GoSub, ShowReminder

return

; Label for reminder
ShowReminder:
    MsgBox, 64, LaTeX to OneNote Ready, Press Ctrl + Alt + L to convert clipboard contents from LaTeX to OneNote.
return

ToggleAutoPaste:
    AutoPaste := !AutoPaste
    Menu, Tray, ToggleCheck, Auto-Paste on Success
return

ExitScript:
    ExitApp
return

; Hotkey: Ctrl + Alt + L
^!l::
    ; Tooltip to indicate activity
    ToolTip, Converting LaTeX to OneNote...
    
    ; Run the script using pythonw
    ; use RunWait to get the exit code
    RunWait, "%VenvPython%" "%ScriptFile%", , Hide UseErrorLevel
    
    ; Check result
    if (ErrorLevel == 0) {
        ToolTip, Success!
        if (AutoPaste) {
            ToolTip, Success! (Pasting...)
            Send, ^v ; Paste
        }
        Sleep, 1500
        ToolTip ; Hide tooltip
    } else {
        ToolTip, Conversion Failed or No LaTeX found.
        Sleep, 2000
        ToolTip ; Hide tooltip
    }
return
