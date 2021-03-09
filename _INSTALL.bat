powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\Start Menu\Programs\Startup\virtual_desktop_enhancer.lnk');$s.TargetPath='%~dp0\virtual_desktop_enhancer.exe';$s.Save()"
