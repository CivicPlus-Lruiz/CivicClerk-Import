@echo off
REM Launch Meeting Import GUI - Modular Version
REM This script launches the Python GUI application

echo.
echo ================================================
echo   Meeting Import Tool v1.6 (Modular)
echo ================================================
echo.
echo Starting application...
echo.

REM Launch the Python script
python main.py

REM Keep window open if there's an error
if errorlevel 1 (
    echo.
    echo ================================================
    echo   Error: Application failed to start
    echo ================================================
    echo.
    pause
)
