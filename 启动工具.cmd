@echo off
set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
if not exist "%PYTHON_EXE%" (
    echo Python runtime not found: %PYTHON_EXE%
    pause
    exit /b 1
)
set "PYTHONDONTWRITEBYTECODE=1"
start "" "%PYTHON_EXE%" -B "%~dp0start_gui.py"
