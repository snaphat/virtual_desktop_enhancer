#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
DetectHiddenWindows, on
I_Icon = app.ico
IfExist, %I_Icon%
    Menu, Tray, Icon, %I_Icon%

; initiate splash on start.
GuiSplash()

; setup key hooks...

; move to left desktop with splash.
^#Left:: ; alt-left.
    Send {LWin down}{Ctrl down}{Left}{Ctrl up}{LWin up}
    id := CurDesktop() - 1
    if (id > 0) ; Bounds check.
        GuiSplash()
Return

; move to right desktop with splash.
^#Right:: ; alt-right.
    Send {LWin down}{Ctrl down}{Right}{Ctrl up}{LWin up}
    id := CurDesktop() + 1
    if (id < NumDesktops()+1) ; Bounds check.
        GuiSplash()
Return

; move active window to left desktop (no splash).
!#Left:: ; win-alt-left.
    WinGetTitle, Title, A
    WinGet, hwnd, ID, A
    id := CurDesktop() - 1
    if (id > 0) { ; Bounds check.
        MoveToDesktop(hwnd, id)
    }
Return

; move active window to right desktop (no splash).
!#Right:: ; win-alt-right.
    WinGetTitle, Title, A
    WinGet, hwnd, ID, A
    id := CurDesktop() + 1
    if (id < NumDesktops()+1) { ; Bounds check.
        MoveToDesktop(hwnd, id)
    }
Return

; move active window to left desktop and follow.
^!#Left:: ; ctrl-win-alt-left.
    WinGetTitle, Title, A
    WinGet, hwnd, ID, A
    id := CurDesktop() - 1
    if (id > 0) { ; Bounds check.
        MoveToDesktop(hwnd, id)
        Send {LWin down}{Ctrl down}{Left}{Ctrl up}{LWin up}
        GuiSplash()
        WinActivate, %Title%
    }
Return

; move active window to right desktop and follow.
^!#Right:: ; ctrl-win-alt-right.
    WinGetTitle, Title, A
    WinGet, hwnd, ID, A
    id := CurDesktop() + 1
    if (id < NumDesktops()+1) { ; Bounds check.
        MoveToDesktop(hwnd, id)
        Send {LWin down}{Ctrl down}{Right}{Ctrl up}{LWin up}
        GuiSplash()
        WinActivate, %Title%
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

MoveToDesktop(hwnd, id)
{
    global

    ; Check window IDs (only attempt to move "valid" windows.
    WinGet, dwStyle, Style, ahk_id %hwnd%
    if ((dwStyle&0x08000000) || !(dwStyle&0x10000000)) {
        return false
    }
    WinGet, dwExStyle, ExStyle, ahk_id %hwnd%
    if (dwExStyle & 0x00000080) {
        return false
    }
    WinGetClass, szClass, ahk_id %hwnd%
    if (szClass = "TApplication") {
        return false
    }
    WinGetTitle, title, ahk_id %hwnd%
    if not (title)
        return false

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
    DllCall(MoveViewToDesktop, "ptr", IVirtualDesktopManagerInternal, "Ptr", pView, "UPtr", IVirtualDesktop, "UInt")
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
GuiSplash(){
    Static
    title:="Splash123"
    global DesktopNum ; Splash variable.
    global sec := 0 ; counter for hiding GUI.
    global mutex := CreateMutex() ; Mutex for fixing timing issues with multiple threads.

    ; Lazy create GUI.
    if not WinExist(title) {
        Gui, Color, 0000FF
        Gui, +ToolWindow -Caption +AlwaysOnTop
        Gui, Font, S120 w2000, "Verdana"
        num := CurDesktop()
        Gui, Add, Text, cWhite vDesktopNum, %num%
        Gui, Show, Center NA, %title%
        WinSet, Transparent, 75, %title%
        Gui, Hide
    }

    ; Update Gui display before starting update thread.
    num := CurDesktop()
    GuiControl,,DesktopNum, %num%
    Gui, Show, NA, %title%

    ; Start thread that will setup timer (Don't stall this thread).
    SetTimer, guiHelp, % 10 ; trigger next gui at 10ms interval (only once though).

    ; Thread for starting guiUpdater.
    guiHelp:
        ; Lock for consistency. We don't want to modify the guiUpdater timer if its thread currently running.
        lock(mutex)
        SetTimer, guiHelp, off
        SetTimer, guiUpdater, % 10 ; trigger next gui at 10ms interval.
        unlock(mutex)
        Return

    ; Thread for updating the GUI.
    guiUpdater:
        ; Lock for consistency. We shouldn't be running until the guiHelp thread has setup our timer.
        lock(mutex)
        static num_store := 1 ; store the last known state.
        sec := sec + 10 ; increment the time for hiding the splash.

        ; Check if the desktop number has changed then update (Avoids Visual blinking).
        num := CurDesktop()
        if (num != num_store) {
            sec = 0 ; Reset Counter.
            num_store := num
            GuiControl,,DesktopNum, %num%
            Gui, Show, NA, %title%
        }

        ; Hide the splash after 1000ms.
        if (sec >= 1000) {
            sec = 0 ; Reset counter.
            SetTimer, guiUpdater, off
            Gui, Hide
        }
        unlock(mutex)
        Return
}
