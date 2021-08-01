#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#include lib.ahk

; Add icon.
I_Icon = app.ico
IfExist, %I_Icon%
    Menu, Tray, Icon, %I_Icon%

; Restart script on resolution change to fix tray icon.
; Workaround because autohotkey icon change commands do nothing after resolution changes.
OnMessage(0x7E, "WM_DISPLAYCHANGE")
WM_DISPLAYCHANGE(wParam, lParam)
{
    mSleep(5000) ; takes ~5 seconds for the resolution change to stabilize.
    Reload
}

; Needs to run as admin.
if (not A_IsAdmin)
{
    Run *RunAs "%A_ScriptFullPath%"  ; Requires v1.0.92.01+
    ExitApp
}

;-------------------------------------------------------------------------
; Map hotkeys.
;-------------------------------------------------------------------------
Loop 10 {
    i := A_Index - 1
    HotKey %SwitchShortcut%%i%,             switchToDesktop_%i%
    HotKey %SwitchShortcut%Numpad%i%,       switchToDesktop_%i%
    HotKey %MoveShortcut%%i%,               MoveActiveWindowToDesktop_%i%
    HotKey %MoveShortcut%Numpad%i%,         MoveActiveWindowToDesktop_%i%
    HotKey %moveFollowShortcut%%i%,         MoveActiveWindowToDesktopAndFollow_%i%
    HotKey %moveFollowShortcut%Numpad%i%,   MoveActiveWindowToDesktopAndFollow_%i%
}
HotKey %SwitchShortcut%Left,                switchToDesktop_L
HotKey %SwitchShortcut%Right,               switchToDesktop_R
HotKey %MoveShortcut%Left,                  MoveActiveWindowToDesktop_L
HotKey %MoveShortcut%Right,                 MoveActiveWindowToDesktop_R
HotKey %moveFollowShortcut%Left,            MoveActiveWindowToDesktopAndFollow_L
HotKey %moveFollowShortcut%Right,           MoveActiveWindowToDesktopAndFollow_R

;-------------------------------------------------------------------------
; Remove dead tray icons.
;-------------------------------------------------------------------------
Tray_Refresh()

;-------------------------------------------------------------------------
; initiate splash on start.
;-------------------------------------------------------------------------
ForkGuiSplashLoop()
