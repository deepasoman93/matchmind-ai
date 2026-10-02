@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo First-time setup: creating Python 3.10 environment...
    py -3.10 -m venv .venv
    if errorlevel 1 goto :error

    call ".venv\Scripts\activate.bat"
    python -m pip install --upgrade pip
    if errorlevel 1 goto :error

    pip install -r requirements.txt
    if errorlevel 1 goto :error
) else (
    call ".venv\Scripts\activate.bat"
)

echo Starting MatchMind AI...
python -m streamlit run app.py
goto :end

:error
echo.
echo Setup could not be completed. Confirm that Python 3.10 is installed.
pause

:end
endlocal
