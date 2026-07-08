@echo off
setlocal EnableExtensions

echo Starting CFI DataClean...
echo.

rem This is the folder where this CMD file is located.
set "APP_ROOT=%~dp0"

rem This is the support folder inside the root folder.
set "APP_DIR=%APP_ROOT%_CFIDataClean_AppFiles"

rem This is the actual PowerShell app file.
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
    echo.
    echo Make sure this CMD file is in the root CFI DataClean folder,
    echo not inside _CFIDataClean_AppFiles.
    pause
    exit /b 1
)

if not exist "%PS1%" (
    echo ERROR: The PowerShell app file was not found.
    echo Expected:
    echo %PS1%
    echo.
    echo Make sure ForestInventoryCleaner.ps1 is inside _CFIDataClean_AppFiles.
    pause
    exit /b 1
)

rem Prefer 32-bit PowerShell because Access/ACE drivers are often 32-bit.
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

pushd "%APP_DIR%"

"%PSEXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%PS1%"

echo.
echo CFI DataClean closed or failed to start.
echo If there is an error above, copy that message.
echo.
pause