@echo off
title Building SLAM Jobcard Auto-Filler Standalone Executable...
cd /d "%~dp0"

echo ========================================================
echo   Building SLAM Jobcard Auto-Filler (.exe)
echo ========================================================
echo.

python build_exe.py

echo.
pause
