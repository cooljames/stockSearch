@echo off
cd /d "%~dp0"
echo Starting Stock Search App...

rem Try to reuse the virtual environment from the googleNews project if it exists
if exist "..\googleNews\.venv\Scripts\python.exe" (
    echo Found virtual environment in googleNews directory. Launching...
    "..\googleNews\.venv\Scripts\python.exe" main.py
    goto end
)

rem Try to use local virtual environment
if exist ".venv\Scripts\python.exe" (
    echo Found local virtual environment. Launching...
    ".venv\Scripts\python.exe" main.py
    goto end
)

rem Try to use system python
echo Virtual environment not found. Trying system Python...
python main.py

:end
if %errorlevel% neq 0 (
    echo.
    echo -------------------------------------------------------------
    echo [ERROR] Could not start the program.
    echo Please make sure the following packages are installed:
    echo.
    echo     pip install yfinance matplotlib openpyxl
    echo -------------------------------------------------------------
    echo.
    pause
)
