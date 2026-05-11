@echo off
setlocal

:: bpmn2visio.bat - BPMN to Visio converter launcher
::
:: Usage:
::   Drag and drop a .bpmn file onto this bat file in File Explorer.
::   The converted _com.vsdx will be created in the same folder.
::
::   Or run from command line:
::   bpmn2visio.bat your_file.bpmn

set SCRIPT_DIR=%~dp0

if "%~1"=="" (
    echo bpmn2visio - BPMN to Visio Converter
    echo.
    echo   Drop a .bpmn file onto this bat file to convert it.
    echo.
    pause
    exit /b 1
)

echo ============================================================
echo  bpmn2visio
echo  Input: %~1
echo ============================================================
echo.

python "%SCRIPT_DIR%bpmn2visio.py" "%~1"

echo.
if %errorlevel% equ 0 (
    echo Done. Open the generated .vsdx file in Visio.
) else (
    echo Error occurred. Check the messages above.
)

echo.
pause
endlocal
