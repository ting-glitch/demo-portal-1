@echo off
cd /d "%~dp0"
echo ====================================================
echo      Starting Logistics Operation Platform...
echo ====================================================
echo.
echo [INFO] Moving to backend directory...
cd backend

echo [INFO] Server is launching. Please keep this window open.
echo.

where python >nul 2>nul
if %errorlevel% equ 0 (
    python app.py
) else if exist "C:\Python314\python.exe" (
    C:\Python314\python.exe app.py
) else (
    echo [ERROR] Python was not found in environment PATH or C:\Python314\python.exe!
    pause
)

echo.
echo ----------------------------------------------------
echo [INFO] Service has been terminated.
pause
