@echo off
setlocal EnableExtensions

set "APP_ROOT=%~dp0"
set "APP_DIR=%APP_ROOT%_CFIDataClean_AppFiles"
set "PS1=%APP_DIR%\ForestInventoryCleaner.ps1"

if not exist "%APP_DIR%" (
    echo CFI DataClean could not find the app files folder.
    echo Expected:
    echo %APP_DIR%
    pause
    exit /b 1
)

if not exist "%PS1%" (
    echo CFI DataClean could not find the PowerShell app file.
    echo Expected:
    echo %PS1%
    pause
    exit /b 1
)

set "PSEXE=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PSEXE%" (
    set "PSEXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
)

if not exist "%PSEXE%" (
    echo Windows PowerShell was not found.
    pause
    exit /b 1
)

set "CFI_APP_ROOT=%APP_ROOT%"

rem Try to remove the GitHub/Internet download block from the extracted files.
"%PSEXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "try { Get-ChildItem -LiteralPath $env:CFI_APP_ROOT -Recurse -File | Unblock-File -ErrorAction SilentlyContinue } catch { }"

start "" /d "%APP_DIR%" "%PSEXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%PS1%"

exit /b 0
