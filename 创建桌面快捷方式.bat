@echo off
cd /d "%~dp0"
echo Creating desktop shortcut...

powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$desktop = [Environment]::GetFolderPath('Desktop');" ^
  "$lnk = $ws.CreateShortcut($desktop + '\NIUERA Quote Workbench.lnk');" ^
  "$lnk.TargetPath = '%~dp0start.bat';" ^
  "$lnk.WorkingDirectory = '%~dp0';" ^
  "$lnk.IconLocation = '%SystemRoot%\System32\shell32.dll,13';" ^
  "$lnk.Description = 'NIUERA Quote Workbench (LAN)';" ^
  "$lnk.Save();"

if exist "%USERPROFILE%\Desktop\NIUERA Quote Workbench.lnk" (
    echo.
    echo  [OK] Desktop shortcut created: "NIUERA Quote Workbench"
    echo       Double-click it anytime to start the app (LAN enabled^).
) else (
    echo.
    echo  [FAIL] Could not create it. Manual way:
    echo         right-click start.bat - Send to - Desktop ^(create shortcut^)
)
echo.
pause
