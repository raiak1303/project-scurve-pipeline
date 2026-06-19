@echo off
title Project S-Curve Pipeline
color 1F

echo =====================================================
echo    PROJECT PROGRESS PIPELINE - Starting...
echo =====================================================
echo.

:: Try OneDrive path first, then regular Desktop
if exist "%USERPROFILE%\OneDrive\Desktop\Projectdata\pipeline.py" (
    cd /d "%USERPROFILE%\OneDrive\Desktop\Projectdata"
) else if exist "%USERPROFILE%\Desktop\Projectdata\pipeline.py" (
    cd /d "%USERPROFILE%\Desktop\Projectdata"
) else (
    echo ERROR: Could not find Projectdata folder!
    echo Make sure pipeline.py is inside your Projectdata folder on Desktop.
    pause
    exit
)

echo Found folder: %CD%
echo.
echo Installing required libraries (first time only)...
pip install pandas openpyxl watchdog plotly --quiet
echo.
echo =====================================================
echo  Watching for Excel changes... Keep this open!
echo  Dashboard will open in your browser now.
echo  Press Ctrl+C to stop.
echo =====================================================
echo.

python pipeline.py

echo.
echo Pipeline stopped.
pause