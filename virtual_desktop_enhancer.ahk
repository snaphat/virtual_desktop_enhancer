#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
DetectHiddenWindows, on
I_Icon = app.ico
IfExist, %I_Icon%
    Menu, Tray, Icon, %I_Icon%

if (not A_IsAdmin)
{
    Run *RunAs "%A_ScriptFullPath%"  ; Requires v1.0.92.01+
    ExitApp
}

; initiate splash on start.
gtitle:="Splash123"
gDesktopNum := "" ; Splash variable.
gsec := 0 ; counter for hiding GUI.
gmutex := CreateMutex() ; Mutex for fixing timing issues with multiple threads.
GuiSplash()

; setup key hooks...

; move to left desktop with splash.
^#Left:: ; alt-left.
    id := CurDesktop() - 1
    if (id > 0) { ; Bounds check.
        Send {LWin down}{Ctrl down}{Left down} ; Separated for spurious wakeups.
        GuiSplash()
        WinActivate, A ; Fixes Spurious Wakeups.
        Send {LWin up}{Ctrl up}{Left up}
    }

Return

; move to right desktop with splash.
^#Right:: ; alt-right.
    id := CurDesktop() + 1
    if (id < NumDesktops()+1) { ; Bounds check.
        Send {LWin down}{Ctrl down}{Right down} ; Separated for spurious wakeups.
        GuiSplash()
        WinActivate, A ; Fixes Spurious Wakeups.
        Send {LWin up}{Ctrl up}{Right up}
    }
Return

; move active window to left desktop (no splash).
!#Left:: ; win-alt-left.
    id := CurDesktop() - 1
    if (id > 0) { ; Bounds check.
        WinGet, hwnd, ID, A
        MoveToDesktop(hwnd, id)
        hwnd := WinActive("A")
        SelectNextWindow() ; allows all windows on a desktop be moved easily.
    }
Return

; move active window to right desktop (no splash).
!#Right:: ; win-alt-right.
    id := CurDesktop() + 1
    if (id < NumDesktops()+1) { ; Bounds check.
        WinGet, hwnd, ID, A
        MoveToDesktop(hwnd, id)
        SelectNextWindow() ; allows all windows on a desktop be moved easily.
    }
Return

; move active window to left desktop and follow.
^!#Left:: ; ctrl-win-alt-left.
    id := CurDesktop() - 1
    if (id > 0) { ; Bounds check.
        WinGet, hwnd, ID, A
        MoveToDesktop(hwnd, id)
        Send {LWin down}{Ctrl down}{Left down}
        GuiSplash()
        Send {LWin up}{Ctrl up}{Left up}
        WinActivate, ahk_id %hwnd%
    }
Return

; move active window to right desktop and follow.
^!#Right:: ; ctrl-win-alt-right.
    id := CurDesktop() + 1
    if (id < NumDesktops()+1) { ; Bounds check.
        WinGet, hwnd, ID, A
        MoveToDesktop(hwnd, id)
        Send {LWin down}{Ctrl down}{Right down}
        GuiSplash()
        Send {LWin up}{Ctrl up}{Right up}
        WinActivate, ahk_id %hwnd%
    }
Return

; Fn to get number of desktops (Windows has no exposed APIs).
NumDesktops() {
    RegRead, DesktopList, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops, VirtualDesktopIDs
    return StrLen(DesktopList) // 32 ; UUIDs are 32-bytes so divide the list size by 32.
}

; Fn to get the current desktop (Windows has no exposed APIs).
CurDesktop() {
    ProcessId := DllCall("GetCurrentProcessId", "UInt")
    DllCall("ProcessIdToSessionId", "UInt", ProcessId, "UInt*", SessionId)
    RegRead, CurrentDesktopId, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo\%SessionId%\VirtualDesktops, CurrentVirtualDesktop
    RegRead, DesktopList, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops, VirtualDesktopIDs

    count := StrLen(DesktopList) ; 32  UUIDs are 32-bytes so divide the list size by 32.
    loop %count% ; Loop over list of desktops (IDs are 32 in length)
        if ( SubStr(DesktopList, ((A_Index - 1) * 32) + 1, 32) = CurrentDesktopId) ; compare ID of current desktop to list entry.
            return A_Index
    return 1 ; if no count, assume only 1 desktop.
}

MoveToDesktop(hwnd, id) {
    global

    ; Check window IDs (only attempt to move "valid" windows.)
    if not (IsValidWindow(hwnd))
        return False

    ; Init...
    IServiceProvider := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{6D5140C1-7436-11CE-8034-00AA006009FA}")
    IVirtualDesktopManagerInternal := ComObjQuery(IServiceProvider, "{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}", "{F31574D6-B682-4CDC-BD56-1827860ABEC6}")
    MoveViewToDesktop := vtable(IVirtualDesktopManagerInternal, 4) ; void MoveViewToDesktop(object pView, IVirtualDesktop desktop);
    GetDesktops := vtable(IVirtualDesktopManagerInternal, 7) ; IObjectArray GetDesktops();
    ImmersiveShell := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{00000000-0000-0000-C000-000000000046}")
    IApplicationViewCollection := ComObjQuery(ImmersiveShell,"{1841C6D7-4F9D-42C0-AF41-8747538F10E5}","{1841C6D7-4F9D-42C0-AF41-8747538F10E5}" ) ; 1607-1809
    GetViewForHwnd := vtable(IApplicationViewCollection, 6) ; (IntPtr hwnd, out IApplicationView view);

    ; Do Magic...
    DllCall(GetViewForHwnd, "UPtr", IApplicationViewCollection, "Ptr", hwnd, "Ptr*", pView, "UInt")
    DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")
    VarSetCapacity(vd_GUID, 16)
    DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &vd_GUID)
    DllCall(vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", id-1, "UPtr", &vd_GUID, "UPtrP", IVirtualDesktop, "UInt")
    DllCall(MoveViewToDesktop, "Ptr", IVirtualDesktopManagerInternal, "Ptr", pView, "UPtr", IVirtualDesktop, "UInt")
}

SelectNextWindow() {
    ; Iterate windows in z-order.
    WinGet, hwnd, ID, A
    Loop {
        hwnd := DllCall("GetWindow",uint,hwnd,int,2) ; 2 = GW_HWNDNEXT
        hwnd := hwnd // 1

        if (hwnd == 0) ; Ran out of windows.
            return break

        if not (IsWindowOnCurrentVirtualDesktop(hwnd))
            continue ; Continue if window not on current desktop (or invalid).

        WinActivate, ahk_id %hwnd% ; Activate next z-order window on current virtual desktop.
        break
    }
}

IsValidWindow(hwnd) {
    if (hwnd == 0)
        return False ; not a valid ID.

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

IsWindowOnCurrentVirtualDesktop(hwnd) {
    if not IsValidWindow(hwnd)
        return False ; not a valid Window.

    ; Init...
    IVirtualDesktopManager := ComObjCreate("{AA509086-5CA9-4C25-8F95-589D3C07B48A}", "{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
    IsWindowOnCurrentVirtualDesktop := vtable(IVirtualDesktopManager, 3)

    ; Do Magic...
    VarSetCapacity(val, 4, 0)
    DllCall(IsWindowOnCurrentVirtualDesktop, "Ptr", IVirtualDesktopManager, "Ptr", hwnd, "Ptr" , &val)
    val := NumGet(&val, "BOOL")
    if (val)
        return True
    else
        return False
}

; Fn to index into IObjectArray objects.
vtable(ppv, idx) {
    Return NumGet(NumGet(1*ppv)+A_PtrSize*idx)
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
        Gui, Color, 0000FF
        Gui, +ToolWindow -Caption +AlwaysOnTop
        Gui, Font, S120 w2000, "Verdana"
        num := CurDesktop()
        Gui, Add, Text, cWhite vgDesktopNum, %num%
        Gui, Show, Center NA, %gtitle%
        WinSet, Transparent, 75, %gtitle%
        Gui, Hide
    }

    ; Update Gui display before starting update thread.
    num := CurDesktop()
    GuiControl,,vgDesktopNum, %num%
    Gui, Show, NA, %gtitle%

    ; Start thread that will setup timer (Don't stall this thread).
    SetTimer, guiHelp, % 10 ; trigger next gui at 10ms interval (only once though).

    ; Thread for starting guiUpdater.
    guiHelp:
        ; Lock for consistency. We don't want to modify the guiUpdater timer if its thread currently running.
        lock(mgutex)
        SetTimer, guiHelp, off
        SetTimer, guiUpdater, % 10 ; trigger next gui at 10ms interval.
        unlock(gmutex)
        Return

    ; Thread for updating the GUI.
    guiUpdater:
        ; Lock for consistency. We shouldn't be running until the guiHelp thread has setup our timer.
        lock(gmutex)
        static num_store := 0 ; store the last known state.
        gsec := gsec + 10 ; increment the time for hiding the splash.

        ; Check if the desktop number has changed then update (Avoids Visual blinking).
        num := CurDesktop()
        if (num != num_store) {
            gsec := 0 ; Reset Counter.
            num_store := num
            GuiControl,,gDesktopNum, %num%
            Gui, Show, NA, %gtitle%
        }

        ; Hide the splash after 1000ms.
        if (gsec >= 650) {
            gsec := 0 ; Reset counter.
            num_store := 0
            SetTimer, guiUpdater, off
            Gui, Hide
        }
        unlock(gmutex)
        Return
}
