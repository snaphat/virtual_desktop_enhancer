
# virtual_desktop_enhancer <img src="https://github.com/snaphat/virtual_desktop_enhancer/blob/main/app.ico" width="32" />

## About
Adds QoL improvements to Windows Virtual Desktop Implementation.

- Adds hotkeys for moving windows around virtual desktops.
- Adds a 650 millisecond splash indicating the current desktop.
- Supports desktop looping via hotkeys.
- Simple & Fast unlike other existing solutions.
- Does not cause spurious taskbar alerts unlike other existing solutions.

*Note: Elevated privileges are required to move elevated (admin) windows.*

## Usage

| Hotkey             | Description                                           |
|--------------------|-------------------------------------------------------|
| Ctrl-Win-1         | Switch to virtual desktop 1.                          |
| Ctrl-Win-2         | Switch to virtual desktop 2.                          |
| Ctrl-Win-3         | Switch to virtual desktop 3.                          |
| Ctrl-Win-4         | Switch to virtual desktop 4.                          |
| Ctrl-Win-5         | Switch to virtual desktop 5.                          |
| Ctrl-Win-6         | Switch to virtual desktop 6.                          |
| Ctrl-Win-7         | Switch to virtual desktop 7.                          |
| Ctrl-Win-8         | Switch to virtual desktop 8.                          |
| Ctrl-Win-9         | Switch to virtual desktop 9.                          |
| Ctrl-Win-0         | Switch to virtual desktop 10.                         |
| Ctrl-Win-Left      | Switch to virtual desktop on the left.                |
| Ctrl-Win-Right     | Switch to virtual desktop on the right.               |
| Win-Alt-1          | Move window to virtual desktop 1.                     |
| Win-Alt-2          | Move window to virtual desktop 2.                     |
| Win-Alt-3          | Move window to virtual desktop 3.                     |
| Win-Alt-4          | Move window to virtual desktop 4.                     |
| Win-Alt-5          | Move window to virtual desktop 5.                     |
| Win-Alt-6          | Move window to virtual desktop 6.                     |
| Win-Alt-7          | Move window to virtual desktop 7.                     |
| Win-Alt-8          | Move window to virtual desktop 8.                     |
| Win-Alt-9          | Move window to virtual desktop 9.                     |
| Win-Alt-0          | Move window to virtual desktop 10.                    |
| Win-Alt-Left       | Move window to virtual desktop on the left.           |
| Win-Alt-Right      | Move window to virtual desktop on the right.          |
| Ctrl-Win-Alt-1     | Move window & Switch to virtual desktop 1.            |
| Ctrl-Win-Alt-2     | Move window & Switch to virtual desktop 2.            |
| Ctrl-Win-Alt-3     | Move window & Switch to virtual desktop 3.            |
| Ctrl-Win-Alt-4     | Move window & Switch to virtual desktop 4.            |
| Ctrl-Win-Alt-5     | Move window & Switch to virtual desktop 5.            |
| Ctrl-Win-Alt-6     | Move window & Switch to virtual desktop 6.            |
| Ctrl-Win-Alt-7     | Move window & Switch to virtual desktop 7.            |
| Ctrl-Win-Alt-8     | Move window & Switch to virtual desktop 8.            |
| Ctrl-Win-Alt-9     | Move window & Switch to virtual desktop 9.            |
| Ctrl-Win-Alt-0     | Move window & Switch to virtual desktop 10.           |
| Ctrl-Win-Alt-Left  | Move window & Switch to virtual desktop on the left.  |
| Ctrl-Win-Alt-Right | Move window & Switch to virtual desktop on the right. |

To autostart virtual_desktop_enhancer on login, run:
```
_INSTALL.bat
```
To remove autostart of virtual_desktop_enhancer on login, run:
```
_UNINSTALL.bat
```

## Configuration
virtual_desktop_enhancer supports the ability to change hotkey prefixes and to enable desktop looping using a configuration file. In order to customize the configuration, rename `config.default.ini` to `config.ini` and edit the file as desired. Hotkey settings use the [AHK hotkey format](https://www.autohotkey.com/docs/Hotkeys.htm).
| Setting                                          | Description                                                              |
|--------------------------------------------------|--------------------------------------------------------------------------|
| switch_to_desktop_prefix                         | Prefix key(s) for _Switch to virtual desktop_ hotkeys.                   |
| move_active_window_to_desktop_prefix             | Prefix key(s) for _Move window to the virtual desktop_ hotkeys.          |
| move_active_window_to_desktop_and_follow_prefix  | Prefix key(s) for _Move window & Switch to the virtual desktop_ hotkeys. |
| enable_desktop_looping                           | Configures desktop looping for shortcuts.                                |
| animate_desktop_looping                          | Configures animating of desktop looping.                                 |


## Images
![Desktop 1](https://github.com/snaphat/virtual_desktop_enhancer/raw/assets/1.png)
![Desktop 2](https://github.com/snaphat/virtual_desktop_enhancer/raw/assets/2.png)

## Video
[![Watch the video](https://img.youtube.com/vi/hZiku5DR0j4/maxresdefault.jpg)](https://www.youtube.com/watch?v=hZiku5DR0j4)
