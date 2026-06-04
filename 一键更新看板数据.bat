@echo off
chcp 65001 >nul
title Update Dashboard Data - Drag Excel/CSV here

REM Use the folder where this .bat lives as the deploy directory
set "DEPLOY=%~dp0"
REM Remove trailing backslash
if "%DEPLOY:~-1%"=="\" set "DEPLOY=%DEPLOY:~0,-1%"
set "PS1=%DEPLOY%\update_data.ps1"

if "%~1"=="" (
    echo Drag an Excel or CSV file onto this icon!
    echo Supports: .xlsx .xls .csv
    pause
    exit /b 1
)

if not exist "%~1" (
    echo File not found: %~1
    pause
    exit /b 1
)

echo ========================================
echo  Zhifa QC Dashboard - Data Update
echo ========================================
echo Source: %~1
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" "%~1" "%DEPLOY%"

echo.
pause
