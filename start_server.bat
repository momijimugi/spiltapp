@echo off
setlocal

cd /d "%~dp0"
set "SPLITLAB_PORT=8000"
set "SPLITLAB_URL=http://127.0.0.1:%SPLITLAB_PORT%/index.html"

where py >nul 2>nul
if %errorlevel%==0 (
  set "SPLITLAB_PYTHON=py"
  goto :start
)

where python >nul 2>nul
if %errorlevel%==0 (
  set "SPLITLAB_PYTHON=python"
  goto :start
)

echo.
echo [SPLITLAB] Python was not found.
echo Install Python 3, then run this file again.
echo https://www.python.org/downloads/
echo.
pause
exit /b 1

:start
title SPLITLAB Local Server
echo.
echo ========================================
echo   SPLITLAB Local Server
echo ========================================
echo.
echo URL: %SPLITLAB_URL%
echo.
echo Keep this window open while using the app.
echo Press Ctrl+C to stop the server.
echo.

start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 1; Start-Process '%SPLITLAB_URL%'"
%SPLITLAB_PYTHON% -m http.server %SPLITLAB_PORT% --bind 127.0.0.1

echo.
echo Server stopped.
pause
