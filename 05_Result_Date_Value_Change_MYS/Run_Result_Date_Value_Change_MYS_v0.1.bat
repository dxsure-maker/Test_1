@echo off
cd /d "%~dp0"
set "EXE=dist\Result_Date_Value_Change_MYS_v0.1.exe"
set "EXE_LOCAL=Result_Date_Value_Change_MYS_v0.1.exe"
set "PYW=..\.venv\Scripts\pythonw.exe"
set "PY=..\.venv\Scripts\python.exe"
set "SCRIPT=Result_Date_Value_Change_MYS_v0.1.py"

if exist "%EXE%" (
    start "" "%EXE%"
    exit /b 0
)

if exist "%EXE_LOCAL%" (
    start "" "%EXE_LOCAL%"
    exit /b 0
)

if exist "%PYW%" (
    start "" "%PYW%" "%SCRIPT%"
    exit /b 0
)

if exist "%PY%" (
    start "" "%PY%" "%SCRIPT%"
    exit /b 0
)

echo Executable or Python environment not found.
echo Expected: "%~dp0%EXE%"
echo Expected: "%~dp0%EXE_LOCAL%"
echo Expected: "%~dp0..\.venv\Scripts\pythonw.exe"
echo Expected: "%~dp0..\.venv\Scripts\python.exe"
pause
