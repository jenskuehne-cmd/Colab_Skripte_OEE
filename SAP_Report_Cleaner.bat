@echo off
REM ============================================================
REM SAP Report Cleaner - Windows Starter
REM Doppelklick zum Starten im Explorer
REM ============================================================

cd /d "%~dp0"

echo ============================================================
echo   SAP Report Cleaner
echo ============================================================
echo.

REM Python finden
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set PYTHON_CMD=python
    goto :run
)

where python3 >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set PYTHON_CMD=python3
    goto :run
)

if exist ".venv\Scripts\python.exe" (
    set PYTHON_CMD=.venv\Scripts\python.exe
    goto :run
)

echo Fehler: Python nicht gefunden!
echo Bitte Python 3 installieren: https://www.python.org/downloads/
pause
exit /b 1

:run
echo Python: %PYTHON_CMD%
echo.

%PYTHON_CMD% sap_report_cleaner_gui.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ein Fehler ist aufgetreten.
    pause
)

