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
_SwitchDesktop                   := NumGet(NumGet(_IVirtualDesktopManagerInternal+0)+9*A_PtrSize)
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
    sleep 5000 ; takes ~5 seconds for the resolution change to stabilize.
    Reload
}

; Needs to run as admin.
if (not A_IsAdmin)
{
    Run *RunAs "%A_ScriptFullPath%"  ; Requires v1.0.92.01+
    ExitApp
}

; Remove dead tray icons.
Tray_Refresh()

; initiate splash on start.
gtitle:="Splash123"
gDesktopNum := "" ; Splash variable.
gsec := 0 ; counter for hiding GUI.
gmutex := CreateMutex() ; Mutex for fixing timing issues with multiple threads.
GuiSplash()

; setup key hooks...

; Switch to specific desktop.
!1:: ; alt-1
    SwitchToDesktop(1)
Return
!2:: ; alt-2
    SwitchToDesktop(2)
Return
!3:: ; alt-3
    SwitchToDesktop(3)
Return
!4:: ; alt-4
    SwitchToDesktop(4)
Return
!5:: ; alt-5
    SwitchToDesktop(5)
Return
!6:: ; alt-6
    SwitchToDesktop(6)
Return
!7:: ; alt-7
    SwitchToDesktop(7)
Return
!8:: ; alt-8
    SwitchToDesktop(8)
Return
!9:: ; alt-9
    SwitchToDesktop(9)
Return
!0:: ; alt-0
    SwitchToDesktop(10)
Return

; move to left desktop with splash.
^#Left:: ; ctrl-win-left.
    idx := CurDesktopIdx() - 1
    if (idx > 0) { ; Bounds check.
        GuiSplash() ; Display splash here to reduce delay.
        Send {LWin down}{Ctrl down}{Left down} ; Separated for spurious wakeups.
        WinActivate, A ; Fixes Spurious Wakeups.
        Send {LWin up}{Ctrl up}{Left up}
    }
Return

; move to right desktop with splash.
^#Right:: ; ctrl-win-right.
    idx := CurDesktopIdx() + 1
    if (idx < NumDesktops()+1) { ; Bounds check.
        GuiSplash() ; Display splash here to reduce delay.
        Send {LWin down}{Ctrl down}{Right down} ; Separated for spurious wakeups.
        WinActivate, A ; Fixes Spurious Wakeups.
        Send {LWin up}{Ctrl up}{Right up}
}
Return

; move active window to left desktop (no splash).
!#Left:: ; win-alt-left.
    idx := CurDesktopIdx() - 1
    if (idx > 0) { ; Bounds check.
        WinGet, hwnd, ID, A
        MoveToDesktop(hwnd, idx)
        SelectNextWindow() ; allows all windows on a desktop be moved easily.
    }
Return

; move active window to right desktop (no splash).
!#Right:: ; win-alt-right.
    idx := CurDesktopIdx() + 1
    if (idx < NumDesktops()+1) { ; Bounds check.
        WinGet, hwnd, ID, A
        MoveToDesktop(hwnd, idx)
        SelectNextWindow() ; allows all windows on a desktop be moved easily.
    }
Return

; move active window to left desktop and follow.
^!#Left:: ; ctrl-win-alt-left.
    idx := CurDesktopIdx() - 1
    if (idx > 0) { ; Bounds check.
        WinGet, hwnd, ID, A
        MoveToDesktop(hwnd, idx)
        GuiSplash() ; Display splash here to reduce delay.
        Send {LWin down}{Ctrl down}{Left down}
        WinActivate, ahk_id %hwnd%
        Send {LWin up}{Ctrl up}{Left up}
    }
Return

; move active window to right desktop and follow.
^!#Right:: ; ctrl-win-alt-right.
    idx := CurDesktopIdx() + 1
    if (idx < NumDesktops()+1) { ; Bounds check.
        WinGet, hwnd, ID, A
        MoveToDesktop(hwnd, idx)
        GuiSplash() ; Display splash here to reduce delay.
        Send {LWin down}{Ctrl down}{Right down}
        WinActivate, ahk_id %hwnd%
        Send {LWin up}{Ctrl up}{Right up}
    }
Return

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
    return ptr
}

; Fn to get a virtual desktop from an index.
DesktopPtrFromIdx(index) {
    global
    VarSetCapacity(GUID, 16)
    DllCall(_GetDesktops, "UPtr", _IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")
    DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &GUID)
    DllCall(NumGet(NumGet(IObjectArray+0)+4*A_PtrSize), "UPtr", IObjectArray, "UInt", index-1, "UPtr", &GUID, "UPtrP", ptr, "UInt")  ; IObjectArray::GetAt
    return ptr
}

; Fn to get the current desktop index from ptr.
DesktopIdxFromPtr(ptr) {
    global

    curGUID := DesktopGUIDFromPtr(ptr)
    count := NumDesktops()
    loop %count% {
        GUID := DesktopGUIDFromPtr(DesktopPtrFromIdx(A_Index))
        if (curGUID = GUID)
            return A_Index
    }
    return 1 ; if failed assume first desktop.
}

; Fn to get a GUID from a virtual desktop.
DesktopGUIDFromPtr(ptr) {
    VarSetCapacity(GUID, 16)
    VarSetCapacity(strGUID, (38 + 1) * 2)
    DllCall(NumGet(NumGet(ptr+0)+4*A_PtrSize), "UPtr", ptr, "UPtr", &GUID, "UInt")
    DllCall("Ole32.dll\StringFromGUID2", "UPtr", &GUID, "UPtr", &strGUID, "Int", 38 + 1)
    return StrGet(&strGUID, "UTF-16")
}

; Fn to move a window to the desktop indicated by index.
MoveToDesktop(hwnd, idx) {
    global

    ; Check window IDs (only attempt to move "valid" windows.)
    if not (IsValidWindow(hwnd))
        return False

    ; Do Magic...
    DllCall(_GetViewForHwnd, "UPtr", _IApplicationViewCollection, "Ptr", hwnd, "Ptr*", pView, "UInt")
    DllCall(_MoveViewToDesktop, "Ptr", _IVirtualDesktopManagerInternal, "Ptr", pView, "UPtr", DesktopPtrFromIdx(idx), "UInt")
}

; Fn Switch to desktop indicated by index.
SwitchToDesktop(idx) {
    global
    if (idx < NumDesktops()+1) { ; Bounds check.
        GuiSplash() ; Display splash here to reduce delay.
        WinActivate, ahk_class Shell_TrayWnd ; Fixes Spurious Wakeups
        DllCall(_SwitchDesktop, "Ptr", _IVirtualDesktopManagerInternal, "Ptr", DesktopPtrFromIdx(idx), "UInt")
        WinActivate, A ; Fixes Spurious Wakeups
    }
}

; Fn to select the next highest window in z-order.
SelectNextWindow() {
    ; Iterate windows in z-order.
    foundWindow := false
    WinGet, hwnd, ID, A
    Loop {
        hwnd := DllCall("GetWindow",uint,hwnd,int,2) ; 2 = GW_HWNDNEXT
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

; Fn to check if a window handle is valid.
IsValidWindow(hwnd) {

    if (hwnd == 0)
        return False ; not a valid ID.

    VarSetCapacity(cloaked,4, 0)
    DllCall("dwmapi\DwmGetWindowAttribute" , "Ptr", hwnd ,"UInt", 14, "Ptr", &cloaked, "UInt", 4)
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
    if (szClass = "TApplication")
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
    val := NumGet(&val, "BOOL")
    return val ? true : false
}

; Fn to create a mutex. Thanks AutoHotKey for not having locks but having threading?
CreateMutex() {
    return DllCall("CreateMutex", Ptr, 0, Int, False, Ptr, 0, Ptr)
}

; Fn to lock a mutex. Thanks AutoHotKey for not having locks but having threading?
lock(mutex) { ; LOL safety first.
    DllCall("WaitForSingleObject", Ptr, mutex, Int, -1) ; INFINITE
}

; Fn to unlock a mutex. Thanks AutoHotKey for not having locks but having threading?
unlock(mutex) {
    DllCall("ReleaseMutex", Ptr, mutex)
}

; Splash screen
GuiSplash() {
    global

    ; Lazy create GUI.
    if not (WinExist(gtitle)) {
        Gui, +E0x08000000 ; No-activate style
        Gui, Color, 0000FF
        Gui, +ToolWindow -Caption +AlwaysOnTop
        Gui, Font, S120 w2000, "Verdana"
        _idx := CurDesktopIdx()
        Gui, Add, Text, cWhite vgDesktopNum, %_idx%
        Gui, Show, Center NA, %gtitle%
        WinSet, Transparent, 75, %gtitle%
        Gui, Hide
        return
    }

    ; Update Gui display before starting update thread.
    _idx := CurDesktopIdx()
    GuiControl,,vgDesktopNum, %_idx%
    Gui, Show, NA, %gtitle%

    ; Start thread that will setup timer (Don't stall this thread).
    SetTimer, guiHelp, % 10 ; trigger next gui at 10ms interval (only once though).

    ; Thread for starting guiUpdater.
    guiHelp:
        ; Lock for consistency. We don't want to modify the guiUpdater timer if its thread currently running.
        lock(gmutex)
        SetTimer, guiHelp, off
        SetTimer, guiUpdater, % 10 ; trigger next gui at 10ms interval.
        unlock(gmutex)
        Return

    ; Thread for updating the GUI.
    guiUpdater:
        ; Lock for consistency. We shouldn't be running until the guiHelp thread has setup our timer.
        lock(gmutex)
        static idx_store := 0 ; store the last known state.
        gsec := gsec + 10 ; increment the time for hiding the splash.

        ; Check if the desktop number has changed then update (Avoids Visual blinking).
        _idx := CurDesktopIdx()
        if (_idx != idx_store) {
            gsec := 0 ; Reset Counter.
            idx_store := _idx
            GuiControl,,gDesktopNum, %_idx%
            Gui, Show, NA, %gtitle%
        }

        ; Hide the splash after 1000ms.
        if (gsec >= 650) {
            gsec := 0 ; Reset counter.
            idx_store := 0
            SetTimer, guiUpdater, off
            Gui, Hide
        }
        unlock(gmutex)
        Return
}
