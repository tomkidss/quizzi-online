@echo off
title Quizzi Admin (PowerShell)
cd /d "%~dp0"

echo Starting Quizzi App via PowerShell...
echo This will open the admin login page in 3 seconds.

:: Open the browser asynchronously (wait 3 seconds for server to start)
start "" cmd /c "timeout /t 3 /nobreak >nul && start http://127.0.0.1:5000/login"

:: Open PowerShell, activate venv and run app.py
start powershell -NoExit -Command ".\venv\Scripts\activate; python app.py"

exit
