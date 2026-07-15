@echo off
setlocal

set "APPDIR=%~dp0"
set "SCRIPT=%APPDIR%\ForestInventoryCleaner.ps1"
set "PS32=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PS32%" (
    set "PS32=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
)

if not exist "%SCRIPT%" (
    echo DADA could not find its app files.
    echo.
    echo Expected file:
    echo "%SCRIPT%"
    echo.
    echo If you opened DADA from a ZIP/compressed folder, click Extract All first,
    echo then run _DADA_AppFiles\DADA-debug.cmd from the extracted folder.
    echo.
    pause
    exit /b 1
)

"%PS32%" -NoProfile -STA -ExecutionPolicy Bypass -File "%SCRIPT%"
set "APP_EXIT=%ERRORLEVEL%"

if not "%APP_EXIT%"=="0" (
    echo.
    echo Database Dad closed because of an error.
    echo The error message above should explain what happened.
    echo.
    pause
)

endlocal
exit /b %APP_EXIT%
