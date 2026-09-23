@echo off
title SLAM Jobcard Auto-Filler
cd /d "%~dp0"

if exist "dist\SLAM_Auto_Filler.exe" (
    start "" "dist\SLAM_Auto_Filler.exe"
    exit /b
)

if exist "SLAM_Auto_Filler.exe" (
    start "" "SLAM_Auto_Filler.exe"
    exit /b
)

python main.py
if errorlevel 1 (
    echo.
    echo An error occurred while running SLAM Auto-Filler.
    pause
)
