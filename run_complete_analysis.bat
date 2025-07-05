@echo off
chcp 65001 >nul
echo ========================================
echo    Complete Excel Analysis Workflow
echo ========================================
echo.

if "%~1"=="" (
    echo Usage: Drag and drop an Excel file onto this batch file
    echo.
    echo Or run: run_complete_analysis.bat "path\to\your\file.xlsx"
    echo.
    echo Example:
    echo   run_complete_analysis.bat "Editing Production Schedule - MW Version.xlsx"
    echo.
    pause
    exit /b 1
)

echo Analyzing: %~1
echo.

python complete_analysis_simple.py "%~1"

echo.
echo Analysis complete! Check the generated files.
pause 