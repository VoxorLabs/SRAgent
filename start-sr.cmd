@echo off
title Voxor Speaker Ready Agent
cd /d "%~dp0"
echo Starting Voxor Speaker Ready Micro-Agent...
echo.
node agent.js
pause
