@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "BACKUP_SCRIPT=%SCRIPT_DIR%outlook_backup.py"

REM ── Try finding python in PATH first ──────────────────────────────────────
where python >nul 2>&1
if %ERRORLEVEL% == 0 (
    python "%BACKUP_SCRIPT%"
    goto :done
)

REM ── Fallback: try py launcher (installed with Python on Windows) ──────────
where py >nul 2>&1
if %ERRORLEVEL% == 0 (
    py "%BACKUP_SCRIPT%"
    goto :done
)

REM ── Fallback: known Python install paths ─────────────────────────────────
for %%V in (313 314 312 311 310 39) do (
    if exist "C:\Users\T0415IL\AppData\Local\Programs\Python\Python%%V\python.exe" (
        "C:\Users\T0415IL\AppData\Local\Programs\Python\Python%%V\python.exe" "%BACKUP_SCRIPT%"
        goto :done
    )
)

REM ── Nothing worked ────────────────────────────────────────────────────────
echo.
echo [!] Python not found on this machine.
echo     This tool requires Python to be installed.
echo     Please contact your administrator.
echo.
pause
exit /b 1

:done
endlocal
