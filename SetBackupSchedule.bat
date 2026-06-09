@echo off
setlocal enabledelayedexpansion

cls
echo ===================================================
echo   STELLANTIS EMAIL ARCHIVE - SCHEDULER
echo ===================================================
echo.

REM ── If time passed as argument (from Python popup), skip prompt ──────────
if not "%~1"=="" (
    set "user_time=%~1"
    goto :PROCESS
)

:GETTIME
echo Please enter time like: 10:30 AM
set /p "user_time=Enter Time: "

:PROCESS
:: Validate AM/PM
echo %user_time% | findstr /i "AM PM" >nul
if %errorlevel% neq 0 (
    echo [!] Invalid format. Use AM or PM.
    goto :GETTIME
)

:: Split time and AM/PM
for /f "tokens=1,2 delims= " %%A in ("%user_time%") do (
    set time_part=%%A
    set ampm=%%B
)

:: Split HH and MM
for /f "tokens=1,2 delims=:" %%A in ("%time_part%") do (
    set hh=%%A
    set mm=%%B
)

:: Basic validation
if "!hh!"=="" goto :GETTIME
if "!mm!"=="" goto :GETTIME

:: Remove leading zeros safely
set /a hh=1%hh%-100
set /a mm=1%mm%-100

:: Convert to 24-hour format
if /I "!ampm!"=="PM" (
    if not "!hh!"=="12" set /a hh=!hh!+12
)

if /I "!ampm!"=="AM" (
    if "!hh!"=="12" set hh=0
)

:: Format properly (always 2 digits)
if !hh! LSS 10 (set hh=0!hh!)
if !mm! LSS 10 (set mm=0!mm!)

set final_time=!hh!:!mm!

echo.
echo Creating task at !final_time! (Mon-Fri)...

:: Create scheduled task
schtasks /create ^
/tn "Stellantis_Email_Backup" ^
/tr "cmd /c cd /d \"%cd%\" && run_backup.bat" ^
/sc weekly ^
/d MON,TUE,WED,THU,FRI ^
/st !final_time! ^
/f ^
/it

echo.
echo RESULT CODE: %errorlevel%

if %errorlevel% equ 0 (
    echo.
    echo Removing AC power restriction...
    powershell -NoProfile -Command "$s=New-Object -ComObject Schedule.Service; $s.Connect(); $t=$s.GetFolder('\').GetTask('Stellantis_Email_Backup'); $d=$t.Definition; $d.Settings.DisallowStartIfOnBatteries=$false; $d.Settings.StopIfGoingOnBatteries=$false; $s.GetFolder('\').RegisterTaskDefinition('Stellantis_Email_Backup',$d,4,'','',3)"
    echo.
    echo ===================================================
    echo SUCCESS: Backup scheduled at %user_time%
    echo Runs on battery AND AC power.
    echo ===================================================
) else (
    echo.
    echo [!] FAILED to create task.
    echo Try running this file as Administrator.
)

echo.
echo Checking if task exists...

schtasks /query /tn "Stellantis_Email_Backup" >nul 2>&1

if %errorlevel% equ 0 (
    echo Task verified in Task Scheduler.
) else (
    echo Task NOT found. Something failed.
)

pause