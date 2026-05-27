@echo off
cd /d "%~dp0"

REM חישוב תאריך היום בפורמט YYYY-MM-DD דרך PowerShell (לא תלוי ב-locale של Windows)
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set TODAY=%%i

if not exist "%~dp0%TODAY%" (
    echo.
    echo שגיאה: לא נמצאה תיקייה לתאריך היום: %TODAY%
    echo צפוי: %~dp0%TODAY%
    echo.
    pause
    exit /b 1
)

python run.py "%~dp0%TODAY%"
pause
