:: harrietobrien

@echo off
setlocal EnableExtensions EnableDelayedExpansion

:: Timestamp for backup
for /f %%I in ('powershell -NoProfile -Command "[DateTime]::Now.ToString(\"yyyyMMdd_HHmmss\")"') do set TS=%%I
set "BACKUP_DIR=%USERPROFILE%\Desktop\OutlookCacheBackup_%TS%"
mkdir "%BACKUP_DIR%" >nul 2>&1

echo Closing Outlook processes ...
for %%P in (outlook.exe hxoutlook.exe hxtsr.exe) do taskkill /f /im %%P >nul 2>&1
for %%P in (SearchProtocolHost.exe OfficeClickToRun.exe FileCoAuth.exe) do taskkill /f /im %%P >nul 2>&1
timeout /t 2 /nobreak >nul

echo.
echo New Outlook UWP cache ...
set "FOUND_UWP=0"
for /d %%D in ("%LOCALAPPDATA%\Packages\Microsoft.OutlookForWindows_*") do (
  if exist "%%D\LocalCache" (
    set "FOUND_UWP=1"
    echo Deleting "%%D\LocalCache"
    rmdir /s /q "%%D\LocalCache"
  )
)
if "!FOUND_UWP!"=="0" echo Not found: %LOCALAPPDATA%\Packages\Microsoft.OutlookForWindows_*\LocalCache

echo.
echo Classic Outlook cache ...
set "CLASSIC1=%LOCALAPPDATA%\Microsoft\Outlook"
set "CLASSIC2=%USERPROFILE%\AppData\Local\Microsoft\Outlook"

:: Resolve absolute paths and de-dup
for %%A in ("%CLASSIC1%") do set "C1=%%~fA"
for %%A in ("%CLASSIC2%") do set "C2=%%~fA"
if /i "%C1%"=="%C2%" set "C2="

if defined C1 call :processClassic "%C1%"
if defined C2 call :processClassic "%C2%"

echo.
echo New Outlook cache deleted (rebuilds on reboot)
echo Classic cache items backed up to: "%BACKUP_DIR%"
echo.
pause
endlocal
goto :eof

:processClassic
set "DIR=%~1"
if not exist "%DIR%" (
  echo Skipped: %DIR% not found
  goto :eof
)

echo Scanning %DIR%
pushd "%DIR%" || goto :afterPushd

:: files
for %%F in (*.ost *.oab *.dat *.tmp *.syd) do (
  if exist "%%F" (
    move /y "%%F" "%BACKUP_DIR%" >nul
    if errorlevel 1 (
      echo   LOCKED in use -> %%F
    ) else (
      echo   Backed up and removed %%F
    )
  )
)

:: folder
if exist "Offline Address Book" (
  move /y "Offline Address Book" "%BACKUP_DIR%\Offline Address Book" >nul
  if errorlevel 1 (
    echo   LOCKED in use -> Offline Address Book
  ) else (
    echo   Backed up and removed Offline Address Book
  )
)

:afterPushd
popd
goto :eof
