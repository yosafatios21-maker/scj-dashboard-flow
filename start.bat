@echo off
cd /d "%~dp0"
echo Installing dependencies...
call npm install
echo Starting SCJ Dashboard...
npm start
pause
