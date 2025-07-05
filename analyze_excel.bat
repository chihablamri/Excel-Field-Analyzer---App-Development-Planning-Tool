@echo off
echo ========================================
echo    Excel Field Analyzer
echo ========================================
echo.

if "%~1"=="" (
    echo Usage: Drag and drop an Excel file onto this batch file
    echo.
    echo Or run: analyze_excel.bat "path\to\your\file.xlsx"
    echo.
    pause
    exit /b 1
)

echo Analyzing: %~1
echo.

python excel_analyzer_cli_simple.py "%~1"

echo.
echo Analysis complete! Check the generated files.
pause 