@echo off
cd /d "%~dp0"
echo ====================================================
echo      GitHub Auto-Push Automation Script
echo ====================================================
echo.

set GIT_CMD=git
set REMOTE_URL=https://github.com/ting-glitch/ops-platform-demo.git

:: Check if git is available in standard environment path
where git >nul 2>nul
if %errorlevel% equ 0 (
    echo [INFO] Git is detected in system environment PATH.
    goto :RunGit
)

:: Check common absolute paths if not found in PATH
if exist "C:\Program Files\Git\cmd\git.exe" (
    set GIT_CMD="C:\Program Files\Git\cmd\git.exe"
    echo [INFO] Git is detected at C:\Program Files\Git\cmd\git.exe.
    goto :RunGit
)
if exist "C:\Program Files\Git\bin\git.exe" (
    set GIT_CMD="C:\Program Files\Git\bin\git.exe"
    echo [INFO] Git is detected at C:\Program Files\Git\bin\git.exe.
    goto :RunGit
)
if exist "%USERPROFILE%\AppData\Local\Programs\Git\cmd\git.exe" (
    set GIT_CMD="%USERPROFILE%\AppData\Local\Programs\Git\cmd\git.exe"
    echo [INFO] Git is detected in user AppData directory.
    goto :RunGit
)
if exist "C:\Program Files (x86)\Git\cmd\git.exe" (
    set GIT_CMD="C:\Program Files (x86)\Git\cmd\git.exe"
    echo [INFO] Git is detected at C:\Program Files (x86)\Git\cmd\git.exe.
    goto :RunGit
)

:: Git is not detected
echo [WARNING] Git is not installed or detected on this machine.
echo [INFO] To push your code, please download and install Git for Windows:
echo        https://git-scm.com/download/win
echo.
echo Once Git is installed, you can double-click this script again to push automatically!
echo.
pause
exit /b

:RunGit
echo [INFO] Initializing Git repository locally...
%GIT_CMD% init

echo [INFO] Configuring remote origin URL...
:: Remove origin if it already exists, then add the correct one
%GIT_CMD% remote remove origin >nul 2>nul
%GIT_CMD% remote add origin %REMOTE_URL%

echo [INFO] Staging files (.gitignore will protect sensitive databases/configs)...
%GIT_CMD% add .

echo [INFO] Committing changes...
%GIT_CMD% commit -m "Initialize Logistics Operations Platform: English settings translation, mobile-friendly CSS UI layout, and Git environment templates"

echo [INFO] Renaming default branch to main...
%GIT_CMD% branch -M main

echo.
echo ====================================================
echo      Pushing to GitHub remote repository...
echo      Remote: %REMOTE_URL%
echo ====================================================
echo.
echo [IMPORTANT] If this is your first time pushing, a secure GitHub login popup
echo             may appear. Please authenticate to complete the push.
echo.

%GIT_CMD% push -u origin main --force

if %errorlevel% equ 0 (
    echo.
    echo ====================================================
    echo [SUCCESS] Your project has been successfully backed up and pushed to GitHub!
    echo           Url: %REMOTE_URL%
    echo ====================================================
) else (
    echo.
    echo ====================================================
    echo [WARNING] Git push encountered an error or was cancelled.
    echo           Please ensure:
    echo           1. You have push access to: %REMOTE_URL%
    echo           2. You completed the login authentication popup.
    echo ====================================================
)

echo.
pause
