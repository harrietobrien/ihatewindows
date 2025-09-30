:: harrietobrien

@echo off
setlocal EnableExtensions EnableDelayedExpansion
goto :main

:try_del_dir
:: %1 = dirpath, %2 = label for echo
set "_TARGET=%~1"
set "_LABEL=%~2"
set "_TRIES=0"
if not exist "%_TARGET%" (
  echo Not found: "%_TARGET%"
  goto :eof
)
:retry_del
set /a _TRIES+=1
echo [%_LABEL%] Attempt %_TRIES% delete "%_TARGET%"
attrib -r -s -h "%_TARGET%" /s /d >nul 2>&1
rmdir /s /q "%_TARGET%" >nul 2>&1
if exist "%_TARGET%" (
  if %_TRIES% lss 3 (
    echo still locked; waiting and retrying...
    timeout /t 2 /nobreak >nul
    goto :retry_del
  ) else (
    echo FAILED to delete: "%_TARGET%"
  )
) else (
  echo deleted.
)
goto :eof

:main
echo Closing Teams + WebView2 processes...
taskkill /f /t /im Teams.exe >nul 2>&1
taskkill /f /t /im ms-teams.exe >nul 2>&1
taskkill /f /t /im msedgewebview2.exe >nul 2>&1

echo.
set "NEWPKGROOT=%LOCALAPPDATA%\Packages"
set "PKGCOUNT=0"

for /d %%D in ("%NEWPKGROOT%\MSTeams_*") do (
  set /a PKGCOUNT+=1
  echo Package [!PKGCOUNT!]: "%%~nxD"

  set "NEWCACHE=%%D\LocalCache"
  set "NEWSTATE_MSTEAMS=%%D\LocalState\Microsoft\MSTeams"

  call :try_del_dir "!NEWCACHE!" "LocalCache"
  call :try_del_dir "!NEWSTATE_MSTEAMS!" "LocalState\\Microsoft\\MSTeams"
)

if %PKGCOUNT%==0 (
  echo No MSTeams_* packages found under "%NEWPKGROOT%"
)

echo.
set "CLASSIC_DIR=%APPDATA%\Microsoft\Teams"
set "CLASSIC_LOCAL=%LOCALAPPDATA%\Microsoft\Teams"
echo Nuke Classic Teams . . .
echo.
if exist "%CLASSIC_DIR%" (
  echo "%CLASSIC_DIR%"
  pushd "%CLASSIC_DIR%"
  for %%D in (
    "Application Cache\Cache"
    "blob_storage"
    "databases"
    "GPUCache"
    "IndexedDB"
    "Local Storage"
    "tmp"
    "Code Cache"
    "Cache"
    "Service Worker\CacheStorage"
    "Service Worker\ScriptCache"
  ) do (
    call :try_del_dir "%%~D" "ClassicProfile"
  )
  for %%F in (*.ldb *.log *.sqlite *.db *.tmp) do (
    if exist "%%F" (
      attrib -r -s -h "%%F" >nul 2>&1
      del /q "%%F" >nul 2>&1
    )
  )
  popd
) else (
  echo Classic Teams profile not found: "%CLASSIC_DIR%"
)
echo.
echo Nuke New Teams . . .
echo.
if exist "%CLASSIC_LOCAL%" (
  echo "%CLASSIC_LOCAL%"
  pushd "%CLASSIC_LOCAL%"
  for %%D in ("GPUCache" "Code Cache" "Cache" "tmp") do (
    call :try_del_dir "%%~D" "ClassicLocal"
  )
  popd
) else (
  echo Skipped: "%CLASSIC_LOCAL%"
)

echo.
echo New + Classic Teams caches deleted (rebuild on reboot)
echo.
pause
endlocal
