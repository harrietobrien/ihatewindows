@echo off
title Clear PyCharm Cache and Config

setlocal

set "p1=%LOCALAPPDATA%\JetBrains\PyCharm*"
set "p2=%APPDATA%\JetBrains\PyCharm*"
set "p3=%USERPROFILE%\.PyCharm*"

for %%P in ("%p1%" "%p2%" "%p3%") do (
    echo Deleting %%~P
    if exist %%~P (
        rmdir /S /Q "%%~P"
    ) else (
        echo Not found: %%~P
    )
)

echo.
echo PyCharm cache and settings cleared.
echo Restart PyCharm to rebuild indexes.
echo.
pause
