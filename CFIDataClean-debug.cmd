@echo off
setlocal EnableExtensions

echo Starting CFI DataClean debug launcher...
echo.

set "APP_ROOT=%~dp0"
set "APP_DIR=%APP_ROOT%_CFIDataClean_AppFiles"
set "PS1=%APP_DIR%\ForestInventoryCleaner.ps1"

echo Root folder:
echo %APP_ROOT%
echo.

echo App files folder:
echo %APP_DIR%
echo.

echo PowerShell script:
echo %PS1%
echo.

if not exist "%APP_DIR%" (
    echo ERROR: The app files folder was not found.
    echo Expected:
    echo %APP_DIR%
    pause
    exit /b 1
)

if not exist "%PS1%" (
    echo ERROR: The PowerShell app file was not found.
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
    echo ERROR: Windows PowerShell was not found.
    pause
    exit /b 1
)

echo Using PowerShell:
echo %PSEXE%
echo.

set "CFI_APP_ROOT=%APP_ROOT%"

echo Removing GitHub/Internet download block if present...
"%PSEXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "try { Get-ChildItem -LiteralPath $env:CFI_APP_ROOT -Recurse -File | Unblock-File -ErrorAction SilentlyContinue; Write-Host 'Unblock step completed.' } catch { Write-Host $_ }"
echo.

echo Current PowerShell execution policies:
"%PSEXE%" -NoLogo -NoProfile -Command "Get-ExecutionPolicy -List"
echo.

echo Launching app...
echo.

pushd "%APP_DIR%"

"%PSEXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%PS1%"

echo.
echo CFI DataClean closed or failed to start.
echo If there is an error above, copy that message.
echo.
pause
