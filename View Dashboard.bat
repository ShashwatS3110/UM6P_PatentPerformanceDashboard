@echo off
REM ============================================================
REM  Double-click to open the dashboard in your web browser.
REM  (If it says the file is missing, run "Refresh Dashboard" first.)
REM ============================================================
if exist "%~dp0UM6P_Dashboard.html" (
  start "" "%~dp0UM6P_Dashboard.html"
) else (
  echo UM6P_Dashboard.html not found.
  echo Please run "Refresh Dashboard.bat" first to generate it.
  pause
)
