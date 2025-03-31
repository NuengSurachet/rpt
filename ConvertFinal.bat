@echo off
echo RPT to Excel Converter (Final Version)
echo ====================================

REM Create folders if they don't exist
if not exist "%~dp0rpt" mkdir "%~dp0rpt"
if not exist "%~dp0excel" mkdir "%~dp0excel"

if "%~1"=="" (
    echo Drag and drop an RPT file onto this batch file to convert it to Excel.
    echo Or run: ConvertFinal.bat path\to\file.rpt
    goto end
)

echo Converting: %~1
powershell -ExecutionPolicy Bypass -File "%~dp0ConvertFinal.ps1" "%~1"
echo.
echo Output will be saved in the 'excel' folder.

:end
echo.
pause
