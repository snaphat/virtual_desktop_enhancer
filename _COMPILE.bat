del virtual_desktop_enhancer.zip virtual_desktop_enhancer.exe
Ahk2Exe.exe /in "virtual_desktop_enhancer.ahk" /icon "app.ico" /compress 2
:loop
@tasklist /fi "imagename eq Ahk2Exe.exe" |find ":" > nul
@if errorlevel 1 goto loop > nul
zip -r  virtual_desktop_enhancer.zip virtual_desktop_enhancer.exe config.default.ini _INSTALL.bat _UNINSTALL.bat README.md VERSION
