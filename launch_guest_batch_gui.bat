@echo off
setlocal EnableExtensions

cd /d "%~dp0"
chcp 65001 >nul
title Order Query GUI Launcher
set PYTHONIOENCODING=utf-8

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python 3 was not found in PATH.
    pause
    exit /b 1
)

python -c "import requests, openpyxl, tkinter" >nul 2>nul
if errorlevel 1 (
    echo [INFO] Installing missing packages: requests openpyxl
    python -m pip install requests openpyxl
    if errorlevel 1 (
        echo [ERROR] Package install failed.
        echo Run: python -m pip install requests openpyxl
        pause
        exit /b 1
    )
)

where pythonw >nul 2>nul
if errorlevel 1 (
    python nike_guest_batch_gui_launcher.pyw
) else (
    start "" /B pythonw nike_guest_batch_gui_launcher.pyw
)

exit /b 0
