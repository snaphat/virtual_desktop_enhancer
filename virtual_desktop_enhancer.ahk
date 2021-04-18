#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
DetectHiddenWindows, on

; Used interfaces.
_IVirtualDesktopManager          := ComObjCreate("{AA509086-5CA9-4C25-8F95-589D3C07B48A}", "{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
_IsWindowOnCurrentVirtualDesktop := NumGet(NumGet(_IVirtualDesktopManager+0)+3*A_PtrSize)
_IServiceProvider                := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{6D5140C1-7436-11CE-8034-00AA006009FA}")
_IVirtualDesktopManagerInternal  := ComObjQuery(_IServiceProvider, "{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}", "{F31574D6-B682-4CDC-BD56-1827860ABEC6}")
_GetDesktopCount                 := NumGet(NumGet(_IVirtualDesktopManagerInternal+0)+3*A_PtrSize)
_MoveViewToDesktop               := NumGet(NumGet(_IVirtualDesktopManagerInternal+0)+4*A_PtrSize)
_GetCurrentDesktop               := NumGet(NumGet(_IVirtualDesktopManagerInternal+0)+6*A_PtrSize)
_GetDesktops                     := NumGet(NumGet(_IVirtualDesktopManagerInternal+0)+7*A_PtrSize)
_ImmersiveShell                  := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{00000000-0000-0000-C000-000000000046}")
_IApplicationViewCollection      := ComObjQuery(_ImmersiveShell,"{1841C6D7-4F9D-42C0-AF41-8747538F10E5}","{1841C6D7-4F9D-42C0-AF41-8747538F10E5}" )
_GetViewForHwnd                  := NumGet(NumGet(_IApplicationViewCollection+0)+6*A_PtrSize)

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
; Read config options in.
;-------------------------------------------------------------------------
IniRead, switchShortcut,     config.ini, config, switch_to_desktop_prefix,                        ^#  ; Default: ctrl-win.
IniRead, moveShortcut,       config.ini, config, move_active_window_to_desktop_prefix,            !#  ; Default: win-alt.
IniRead, moveFollowShortcut, config.ini, config, move_active_window_to_desktop_and_follow_prefix, ^!# ; Default: ctrl-win-alt.

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

; setup key hooks...

;-------------------------------------------------------------------------
; switch to specific desktop (non-animated) or
; switch tp left or right desktop (animated).
;-------------------------------------------------------------------------
SwitchToDesktop_1() {
    SwitchToDesktop(1)
}
SwitchToDesktop_2() {
    SwitchToDesktop(2)
}
SwitchToDesktop_3() {
    SwitchToDesktop(3)
}
SwitchToDesktop_4() {
    SwitchToDesktop(4)
}
SwitchToDesktop_5() {
    SwitchToDesktop(5)
}
SwitchToDesktop_6() {
    SwitchToDesktop(6)
}
SwitchToDesktop_7() {
    SwitchToDesktop(7)
}
SwitchToDesktop_8() {
    SwitchToDesktop(8)
}
SwitchToDesktop_9() {
    SwitchToDesktop(9)
}
SwitchToDesktop_0() {
    SwitchToDesktop(10)
}
switchToDesktop_L() {
    idx := CurDesktopIdx() - 1
    SwitchToDesktop(idx)
}
switchToDesktop_R() {
    idx := CurDesktopIdx() + 1
    SwitchToDesktop(idx)
}

;-------------------------------------------------------------------------
; move active window to specific desktop (no splash) or
; move active window to left desktop (no splash).
;-------------------------------------------------------------------------
MoveActiveWindowToDesktop_1() {
    MoveActiveWindowToDesktop(1)
    SelectNextWindow(1)
}
MoveActiveWindowToDesktop_2() {
    MoveActiveWindowToDesktop(2)
    SelectNextWindow(2)
}
MoveActiveWindowToDesktop_3() {
    MoveActiveWindowToDesktop(3)
    SelectNextWindow(3)
}
MoveActiveWindowToDesktop_4() {
    MoveActiveWindowToDesktop(4)
    SelectNextWindow(4)
}
MoveActiveWindowToDesktop_5() {
    MoveActiveWindowToDesktop(5)
    SelectNextWindow(5)
}
MoveActiveWindowToDesktop_6() {
    MoveActiveWindowToDesktop(6)
    SelectNextWindow(6)
}
MoveActiveWindowToDesktop_7() {
    MoveActiveWindowToDesktop(7)
    SelectNextWindow(7)
}
MoveActiveWindowToDesktop_8() {
    MoveActiveWindowToDesktop(8)
    SelectNextWindow(8)
}
MoveActiveWindowToDesktop_9() {
    MoveActiveWindowToDesktop(9)
    SelectNextWindow(9)
}
MoveActiveWindowToDesktop_0() {
    MoveActiveWindowToDesktop(10)
    SelectNextWindow(10)
}
MoveActiveWindowToDesktop_L() {
    idx := CurDesktopIdx() - 1
    MoveActiveWindowToDesktop(idx)
    SelectNextWindow(idx)
}
MoveActiveWindowToDesktop_R() {
    idx := CurDesktopIdx() + 1
    MoveActiveWindowToDesktop(idx)
    SelectNextWindow(idx)
}

;-------------------------------------------------------------------------
; move active window to specific desktop and follow (non-animated) or
; move active window to left desktop and follow (animated).
;-------------------------------------------------------------------------
MoveActiveWindowToDesktopAndFollow_1() {
    MoveActiveWindowToDesktop(1)
    SwitchToDesktop(1)
}
MoveActiveWindowToDesktopAndFollow_2() {
    MoveActiveWindowToDesktop(2)
    SwitchToDesktop(2)
}
MoveActiveWindowToDesktopAndFollow_3() {
    MoveActiveWindowToDesktop(3)
    SwitchToDesktop(3)
}
MoveActiveWindowToDesktopAndFollow_4() {
    MoveActiveWindowToDesktop(4)
    SwitchToDesktop(4)
}
MoveActiveWindowToDesktopAndFollow_5() {
    MoveActiveWindowToDesktop(5)
    SwitchToDesktop(5)
}
MoveActiveWindowToDesktopAndFollow_6() {
    MoveActiveWindowToDesktop(6)
    SwitchToDesktop(6)
}
MoveActiveWindowToDesktopAndFollow_7() {
    MoveActiveWindowToDesktop(7)
    SwitchToDesktop(7)
}
MoveActiveWindowToDesktopAndFollow_8() {
    MoveActiveWindowToDesktop(8)
    SwitchToDesktop(8)
}
MoveActiveWindowToDesktopAndFollow_9() {
    MoveActiveWindowToDesktop(9)
    SwitchToDesktop(9)
}
MoveActiveWindowToDesktopAndFollow_0() {
    MoveActiveWindowToDesktop(10)
    SwitchToDesktop(10)
}
MoveActiveWindowToDesktopAndFollow_L() {
    idx := CurDesktopIdx() - 1
    MoveActiveWindowToDesktop(idx)
    SwitchToDesktop(idx)
}
MoveActiveWindowToDesktopAndFollow_R() {
    idx := CurDesktopIdx() + 1
    MoveActiveWindowToDesktop(idx)
    SwitchToDesktop(idx)
}

;-------------------------------------------------------------------------

; Fn to Remove dead tray icons.
Tray_Refresh() {
    WM_MOUSEMOVE := 0x200
    detectHiddenWin := A_DetectHiddenWindows
    DetectHiddenWindows, On

    allTitles := ["ahk_class Shell_TrayWnd"
            , "ahk_class NotifyIconOverflowWindow"]
    allControls := ["ToolbarWindow321"
                ,"ToolbarWindow322"
                ,"ToolbarWindow323"
                ,"ToolbarWindow324"]
    allIconSizes := [24,32]

    for id, title in allTitles {
        for id, controlName in allControls {
            for id, iconSize in allIconSizes {
                ControlGetPos, xTray,yTray,wdTray,htTray,% controlName,% title
                y := htTray - 10
                While (y > 0) {
                    x := wdTray - iconSize/2
                    While (x > 0) {
                        point := (y << 16) + x
                        PostMessage,% WM_MOUSEMOVE, 0,% point,% controlName,% title
                        x -= iconSize/2
                    }
                    y -= iconSize/2
                }
            }
        }
    }

    DetectHiddenWindows, %detectHiddenWin%
}

; Fn to get number of desktops.
NumDesktops() {
    global

    DllCall(_GetDesktopCount, "Ptr", _IVirtualDesktopManagerInternal, "UInt*", desktopCount)
    if ErrorLevel {
        line := A_LineNumber - 2
        MsgBox,,, Error in function '%A_ThisFunc%' on line %line%!`n`nError: '%A_LastError%'
    }
    return desktopCount
}

; Fn to get the current desktop index.
CurDesktopIdx() {
    return DesktopIdxFromPtr(CurDesktopPtr())
}

; Fn to get the current virtual desktop from an index.
CurDesktopPtr() {
    global

    DllCall(_GetCurrentDesktop, "UPtr", _IVirtualDesktopManagerInternal, "UPtrP", ptr, "Uint")
    if ErrorLevel {
        line := A_LineNumber - 2
        MsgBox,,, Error in function '%A_ThisFunc%' on line %line%!`n`nError: '%A_LastError%'
    }
    return ptr
}

; Fn to get a virtual desktop from an index.
DesktopPtrFromIdx(index) {
    global

    VarSetCapacity(GUID, 16)
    DllCall(_GetDesktops, "UPtr", _IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")
    if ErrorLevel {
        line := A_LineNumber - 2
        MsgBox,,, Error in function '%A_ThisFunc%' on line %line%!`n`nError: '%A_LastError%'
    }
    DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &GUID)
    if ErrorLevel {
        line := A_LineNumber - 2
        MsgBox,,, Error in function '%A_ThisFunc%' on line %line%!`n`nError: '%A_LastError%'
    }
    DllCall(NumGet(NumGet(IObjectArray+0)+4*A_PtrSize), "UPtr", IObjectArray, "UInt", index-1, "UPtr", &GUID, "UPtrP", ptr, "UInt")  ; IObjectArray::GetAt
    if ErrorLevel {
        line := A_LineNumber - 2
        MsgBox,,, Error in function '%A_ThisFunc%' on line %line%!`n`nError: '%A_LastError%'
    }
    ObjRelease(IObjectArray) ; Clear comm object memory.
    return ptr
}

; Fn to get the current desktop index from ptr.
DesktopIdxFromPtr(ptr) {
    global

    curGUID := DesktopGUIDFromPtr(ptr)
    count := NumDesktops()
    loop %count% {
        GUID := DesktopGUIDFromPtr(DesktopPtrFromIdx(A_Index))
        if (curGUID == GUID)
            return A_Index
    }
    return 1 ; if failed assume first desktop.
}

; Fn to get a GUID from a virtual desktop.
DesktopGUIDFromPtr(ptr) {
    VarSetCapacity(GUID, 16)
    VarSetCapacity(strGUID, (38 + 1) * 2)
    DllCall(NumGet(NumGet(ptr+0)+4*A_PtrSize), "UPtr", ptr, "UPtr", &GUID, "UInt")
    if ErrorLevel {
        ; This API can be racey and fail if desktops are removed
        ; so returning -1 in the event of failure is necessary. It is a transient issue
        ; anyway so it won't cause any issues.
        ;line := A_LineNumber - 2
        ;MsgBox,,, Error in function '%A_ThisFunc%' on line %line%!`n`nError: '%A_LastError%'
        return -1
    }
    DllCall("Ole32.dll\StringFromGUID2", "UPtr", &GUID, "UPtr", &strGUID, "Int", 38 + 1)
    if ErrorLevel {
        line := A_LineNumber - 2
        MsgBox,,, Error in function '%A_ThisFunc%' on line %line%!`n`nError: '%A_LastError%'
    }
    return StrGet(&strGUID, "UTF-16")
}

; Fn to move a window to the desktop indicated by index.
MoveActiveWindowToDesktop(idx) {
    global

    if (idx > 0 && idx < NumDesktops()+1 && idx != CurDesktopIdx()) { ; Bounds check.
        ; Check window IDs (only attempt to move "valid" windows.)
        WinGet, hwnd, ID, A
        if not (IsValidWindow(hwnd)) {
            return False
        }

        ; Do Magic...
        DllCall(_GetViewForHwnd, "UPtr", _IApplicationViewCollection, "Ptr", hwnd, "Ptr*", pView, "UInt")
        if ErrorLevel
            MsgBox,,, Error in %A_ThisFunc% at %A_LineNumber%
        DllCall(_MoveViewToDesktop, "Ptr", _IVirtualDesktopManagerInternal, "Ptr", pView, "UPtr", DesktopPtrFromIdx(idx), "UInt")
        if ErrorLevel
            MsgBox,,, Error in %A_ThisFunc% at %A_LineNumber%
    }
}

; Fn Switch to desktop indicated by index.
SwitchToDesktop(idx) {
    if (idx > 0 && idx < NumDesktops()+1  && idx != CurDesktopIdx()) { ; Bounds check.
        diff := idx - CurDesktopIdx()
        if (diff < 0)
            dir := "Left"
        else
            dir := "Right"
        diff := abs(diff)
        loop %diff% {
            Send {LWin down}{Ctrl down}{%dir% down} ; Separated to fix spurious wakeups.
            mSleep(50)
            Send {LWin up}{Ctrl up}{%dir% up}
        }
    }
}

; Fn to select the next highest window in z-order. This allows windows to be moved easily.
SelectNextWindow(idx) {
    ; Iterate windows in z-order.
    if (idx > 0 && idx < NumDesktops()+1 && idx != CurDesktopIdx()) { ; Bounds check.
        foundWindow := false
        WinGet, hwnd, ID, A
        Loop {
            hwnd := DllCall("GetWindow",uint,hwnd,int,2) ; 2 = GW_HWNDNEXT
            if ErrorLevel
                MsgBox,,, Error in %A_ThisFunc% at %A_LineNumber%
            hwnd := hwnd // 1

            if (hwnd == 0)
                break  ; Ran out of windows.

            if not (IsWindowOnCurrentVirtualDesktop(hwnd))
                continue ; Continue if window not on current desktop (or invalid).

            foundWindow := true
            WinActivate, ahk_id %hwnd% ; Activate next z-order window on current virtual desktop.
            break
        }
    }
}

; Fn to check if a window handle is valid.
IsValidWindow(hwnd) {

    if (hwnd == 0)
        return False ; not a valid ID.

    VarSetCapacity(cloaked,4, 0)
    DllCall("dwmapi\DwmGetWindowAttribute" , "Ptr", hwnd ,"UInt", 14, "Ptr", &cloaked, "UInt", 4)
    if ErrorLevel {
        line := A_LineNumber - 2
        MsgBox,,, Error in function '%A_ThisFunc%' on line %line%!`n`nError: '%A_LastError%'
    }
    val := NumGet(cloaked, "UInt") ; DWMWA_CLOAKED value.
    if (val != 0) ; Needed for weeding out Windows10 Apps that are sleeping.
        return False ; Window is Cloaked.

    WinGet, stat, MinMax, ahk_id %hwnd%
    if (stat == -1)
        return False ; iconified so ignore.

    WinGet, dwStyle, Style, ahk_id %hwnd%
    if ((dwStyle&0x08000000) || !(dwStyle&0x10000000))
        return False ; no activate or not-visible.

    WinGet, dwExStyle, ExStyle, ahk_id %hwnd%
    if (dwExStyle & 0x00000080)
        return False ; Tool Window.

    WinGetClass, szClass, ahk_id %hwnd%
    if ((szClass == "TApplication") || (szClass == "Windows.UI.Core.CoreWindow"))
        return False ; Some delphi class window type.

    WinGetTitle, title, ahk_id %hwnd%
    if not (title) ; No title so not valid.
        return False
    return True
}

; Fn to check if a window is on the current virtual desktop.
IsWindowOnCurrentVirtualDesktop(hwnd) {
    global

    if not IsValidWindow(hwnd)
        return False ; not a valid Window.

    ; Do Magic...
    VarSetCapacity(val, 4, 0)
    DllCall(_IsWindowOnCurrentVirtualDesktop, "Ptr", _IVirtualDesktopManager, "Ptr", hwnd, "Ptr" , &val)
    if ErrorLevel {
        line := A_LineNumber - 2
        MsgBox,,, Error in function '%A_ThisFunc%' on line %line%!`n`nError: '%A_LastError%'
    }
    val := NumGet(&val, "BOOL")
    return val ? true : false
}

; Resize Text for GUI.
SetTextAndResize(controlHwnd, newText) {
    Gui 9:Font, S120 w2000, "Verdana"
    Gui 9:Add, Text, cWhite, %newText%
    GuiControlGet T, 9:Pos, Static1
    Gui 9:Destroy

    GuiControl,, %controlHwnd%, %newText%
    GuiControl Move, %controlHwnd%, % "h" TH " w" TW
}

; Accurate sleep.
mSleep(ms) {
    static lazy
    if (lazy != True) {
        DllCall("Winmm.dll\timeBeginPeriod", UInt, 1)
        if ErrorLevel
            MsgBox,,, Error in %A_ThisFunc% at %A_LineNumber%
        lazy := True
    }

    DllCall("Sleep", UInt, ms)
    if ErrorLevel {
        line := A_LineNumber - 2
        MsgBox,,, Error in function '%A_ThisFunc%' on line %line%!`n`nError: '%A_LastError%'
    }
}

; Splash screen
ForkGuiSplashLoop() {
    title:="Splash123"
    global gDesktopNum := "" ; Splash variable.
    Gui, +E0x08000000 ; No-activate style
    Gui, Color, 0000FF
    Gui, +ToolWindow -Caption +AlwaysOnTop
    Gui, Font, S120 w2000, "Verdana"
    Gui, Add, Text, cWhite vgDesktopNum,
    Gui, Show, Center NA, %title%
    WinSet, Transparent, 75, %title%
    Gui, Hide

    ; Loop to update the GUI forever.
    loop {
        mSleep(10)
        static sec := 0 ; counter for hiding GUI.
        static idx_store := 0 ; store the last known state.
        sec := sec + 10 ; increment the time for hiding the splash.

        ; Check if the desktop number has changed then update (Avoids Visual blinking).
        idx := CurDesktopIdx()
        if (idx != idx_store && idx is number) {
            sec := 0 ; Reset Counter.
            idx_store := idx
            GuiControl,,gDesktopNum, %idx%
            GuiControlGet, hwnd, Hwnd, gDesktopNum
            SetTextAndResize(hwnd, idx)
            Gui, Show, NA AutoSize, %title%
        }

        ; Hide the splash after 1000ms.
        if (sec >= 650) {
            sec := 0 ; Reset counter.
            Gui, Hide
        }
    }
}
