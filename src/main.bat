@echo off
:: Check if python is installed
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Python is not installed or not in your PATH.
    echo Please install Python 3.6 or newer from https://www.python.org/downloads/
    pause
    goto end
)

:: Run the analyzer
python ..\dumpsys_analyzer.py
pause
:end
