@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Convert-Invoice.ps1" %*
set "exitcode=%errorlevel%"
echo.
if not "%exitcode%"=="0" echo Conversion failed. See the message above.
if "%exitcode%"=="0" echo Finished. Your files are in the output folder.
echo.
pause
exit /b %exitcode%

