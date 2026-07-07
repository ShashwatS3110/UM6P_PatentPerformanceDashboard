@echo off
REM ============================================================
REM  Double-click this after replacing any of the data files
REM  to rebuild the dashboard with the latest numbers.
REM ============================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\refresh_dashboard.ps1"
echo.
pause
