@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0"
chcp 65001 >nul
title Build NikeGuestBatchGUI EXE

set "APP_NAME=NikeGuestBatchGUI"
set "TARGET_DIR=%CD%\dist\%APP_NAME%"
set "EXE_PATH=%TARGET_DIR%\%APP_NAME%.exe"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python 3 was not found in PATH.
    exit /b 1
)

python -m pip show pyinstaller >nul 2>nul
if errorlevel 1 (
    echo [INFO] Installing PyInstaller...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] PyInstaller install failed.
        exit /b 1
    )
)

python -c "import requests, openpyxl, tkinter" >nul 2>nul
if errorlevel 1 (
    echo [INFO] Installing runtime packages: requests openpyxl
    python -m pip install requests openpyxl
    if errorlevel 1 (
        echo [ERROR] Runtime package install failed.
        exit /b 1
    )
)

echo [INFO] Building %APP_NAME%...
python -m PyInstaller --noconfirm --clean --windowed --onedir --name "%APP_NAME%" nike_guest_batch_gui_launcher.pyw

if errorlevel 1 (
    echo [ERROR] Build failed.
    exit /b 1
)

if not exist "!EXE_PATH!" (
    echo [ERROR] Built EXE not found: !EXE_PATH!
    exit /b 1
)

rem Never copy local settings or proxy credentials into a distribution.

echo.
echo [OK] Build complete.
echo [OK] Double-click this EXE: !EXE_PATH!
echo [OK] Or double-click this launcher: %CD%\launch_guest_batch_gui.bat
