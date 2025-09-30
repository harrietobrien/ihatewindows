:: harrietobrien

@echo off
setlocal EnableExtensions EnableDelayedExpansion

:: Get a safe timestamp via PowerShell (works without admin)
for /f %%I in ('powershell -NoProfile -Command "[DateTime]::Now.ToString(\"yyyyMMdd_HHmmss\")"') do set TS=%%I

set "BACKUP_DIR=%USERPROFILE%\Desktop\OutlookCacheBackup_%TS%"
mkdir "%BACKUP_DIR%" >nul 2>&1

echo Closing Outlook processes...
taskkill /f /im outlook.exe >nul 2>&1
taskkill /f /im hxoutlook.exe >nul 2>&1

echo.
echo New Outlook (UWP) cache . . .
set "NEWO=%LOCALAPPDATA%\Packages\Microsoft.OutlookForWindows_8wekyb3d8bbwe\LocalCache"
if exist "%NEWO%" (
  echo Deleting "%NEWO%"
  rmdir /s /q "%NEWO%"
)
if not exist "%NEWO%" echo Not found: "%NEWO%"

echo.
echo Classic Outlook cache (OST/OAB/Databases) . . . 
set "CLASSIC1=%LOCALAPPDATA%\Microsoft\Outlook"
set "CLASSIC2=%USERPROFILE%\AppData\Local\Microsoft\Outlook"

for %%P in ("%CLASSIC1%" "%CLASSIC2%") do (
  if exist %%~P (
    echo Scanning %%~P
    pushd %%~P
    for %%F in (*.ost *.oab *.dat *.tmp *.syd) do (
      if exist "%%F" (
        echo   Moving "%%F" -> "%BACKUP_DIR%"
        move /y "%%F" "%BACKUP_DIR%" >nul
      )
    )
    if exist "Offline Address Book" (
      echo Moving "Offline Address Book" -> "%BACKUP_DIR%\Offline Address Book"
      move /y "Offline Address Book" "%BACKUP_DIR%\Offline Address Book" >nul
    )
    popd
  )
  if not exist %%~P echo Skipped: %%~P
)

echo.
echo Done.
echo - New Outlook cache deleted (rebuilds on reboot)
echo - Classic cache files moved to: "%BACKUP_DIR%"
echo.
pause
endlocal
