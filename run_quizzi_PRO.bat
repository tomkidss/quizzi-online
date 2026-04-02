@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Quizzi - Auto Start (PRO)
cd /d %~dp0

REM ================== CONFIG ==================
set "PYTHON_EXE=.\venv\Scripts\python.exe"
set "APP_FILE=app.py"
set "PORT=5000"
set "LOG_OUT=server_stdout.log"
set "LOG_ERR=server_stderr.log"
REM ============================================

echo ========================================
echo   Quizzi PRO Auto Start
echo ========================================

REM --- Basic checks ---
if not exist "%APP_FILE%" (
  echo ERROR: "%APP_FILE%" not found in %CD%
  pause
  exit /b 1
)

if not exist "%PYTHON_EXE%" (
  echo ERROR: Python not found at "%PYTHON_EXE%"
  echo       Please check C:\Python314
  pause
  exit /b 1
)

REM =====================================================
REM 0) Resolve LAN IP (prefer the active IPv4 with best metric)
REM =====================================================
set "LAN_IP="
for /f "usebackq delims=" %%I in (`powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ip = (Get-NetIPAddress -AddressFamily IPv4 ^| Where-Object { $_.IPAddress -ne '127.0.0.1' -and $_.IPAddress -notlike '169.254*' -and $_.ValidLifetime -gt 0 } ^| Sort-Object InterfaceMetric ^| Select-Object -First 1 -ExpandProperty IPAddress); if($ip){$ip}else{'127.0.0.1'}"` ) do (
  set "LAN_IP=%%I"
)

if "%LAN_IP%"=="" set "LAN_IP=127.0.0.1"

set "URL_MAIN=http://%LAN_IP%:%PORT%/login"

echo Detected LAN IP: %LAN_IP%
echo Will open: %URL_MAIN%

REM =====================================================
REM 1) Kill anything using PORT (5000) to avoid conflicts
REM =====================================================
echo [1/4] Checking port %PORT%...
for /f "tokens=5" %%P in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  echo    - Port %PORT% is in use. Killing PID %%P ...
  taskkill /F /PID %%P >nul 2>&1
)

REM ============================================
REM 2) Start server hidden/background + capture PID
REM    Start with python.exe (hidden) so we can log stdout/stderr.
REM ============================================
echo [2/4] Starting server in background...
set "SERVER_PID="

REM clear old logs
del /q "%LOG_OUT%" "%LOG_ERR%" >nul 2>&1

start "" /B "%PYTHON_EXE%" "%APP_FILE%" > "%LOG_OUT%" 2> "%LOG_ERR%"

if errorlevel 1 (
  echo ERROR: Could not start server.
  pause
  exit /b 1
)

echo    - Server started in background.

REM ============================================
REM 3) Wait for server to respond, then open ONE page
REM ============================================
echo [3/4] Waiting for server...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$u='%URL_MAIN%'; $ok=$false; for($i=0;$i -lt 40;$i++){ try{ (Invoke-WebRequest -UseBasicParsing -TimeoutSec 1 $u) ^| Out-Null; $ok=$true; break } catch{} Start-Sleep -Milliseconds 300 } if(-not $ok){ exit 1 }" >nul 2>&1

if errorlevel 1 (
  echo WARNING: Server did not respond. Opening page anyway...
  echo         Check logs: %LOG_OUT% and %LOG_ERR%
) else (
  echo    - Server is up.
)

echo Opening: %URL_MAIN%
start "" "%URL_MAIN%"

REM Find the PID of the python process listening on the port
for /f "tokens=5" %%P in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  set "SERVER_PID=%%P"
)

REM ============================================
REM 4) Stop server on key press
REM ============================================
echo.
echo ========================================
echo  Server is running (PID %SERVER_PID%).
echo  Press ANY KEY to stop the server...
echo ========================================
pause >nul

echo Stopping server...
taskkill /F /PID %SERVER_PID% >nul 2>&1

REM Extra safety: kill port again (in case PID changed)
for /f "tokens=5" %%P in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  taskkill /F /PID %%P >nul 2>&1
)

echo Done.
endlocal
exit /b 0
