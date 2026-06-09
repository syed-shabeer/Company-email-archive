@echo off
setlocal

REM ── Move into the folder where this .bat lives ────────────────────────────
cd /d "%~dp0"

set "PORT=8000"

REM ── Check if port 8000 is already in use, try 8001 as fallback ───────────
netstat -ano | findstr ":%PORT% " >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo [~] Port %PORT% is in use, trying 8001...
    set "PORT=8001"
)

REM ── Start the HTTP server from THIS folder ────────────────────────────────
start "Stellantis Email Server" /min cmd /c "cd /d "%~dp0" && python -m http.server %PORT%"

REM ── Start the Explorer helper server on port 9999 ────────────────────────
start "Stellantis Explorer Helper" /min cmd /c "cd /d "%~dp0" && python explorer_helper.py"

REM ── Wait 2 seconds for the server to be ready before opening browser ─────
timeout /t 2 /nobreak >nul

REM ── Open the viewer in the default browser ───────────────────────────────
start "" "http://localhost:%PORT%/search.html"

echo.
echo [+] Email Archive viewer opened at: http://localhost:%PORT%/search.html
echo.
echo     Keep this window open while using the viewer.
echo     Close it when done to stop the server.
echo.
pause

REM ── Kill both servers when user closes this window ──────────────────────
taskkill /fi "WINDOWTITLE eq Stellantis Email Server" /f >nul 2>&1
taskkill /fi "WINDOWTITLE eq Stellantis Explorer Helper" /f >nul 2>&1

endlocal
