@echo off
setlocal
set "INSTDIR=%LOCALAPPDATA%\Renamer"
del "%USERPROFILE%\Desktop\Renamer.lnk" >nul 2>&1
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Renamer.lnk" >nul 2>&1
rmdir /s /q "%INSTDIR%" >nul 2>&1
echo Renamer has been uninstalled.
pause