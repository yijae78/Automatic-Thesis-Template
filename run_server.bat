@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Installing dependencies if needed...
call npm install
echo.
echo Starting server...
node server.js
pause
