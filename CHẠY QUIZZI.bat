@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Quizzi Realtime - Dang khoi dong...
cd /d "%~dp0"

REM ================== CẤU HÌNH ==================
set "PYTHON=.\python\python.exe"
set "PORT=5000"
set "APP=app.py"
REM ================================================

echo.
echo  ==========================================
echo    QUIZZI REALTIME - He thong thi truc tuyen
echo  ==========================================
echo.

REM Kiểm tra Python portable tồn tại
if not exist "%PYTHON%" (
  echo  [LOI] Khong tim thay Python portable tai: %PYTHON%
  echo  Vui long lien he nguoi cung cap phan mem.
  pause
  exit /b 1
)

if not exist "%APP%" (
  echo  [LOI] Khong tim thay file "%APP%"
  pause
  exit /b 1
)

REM =========================================
REM  Phát hiện IP LAN tự động
REM =========================================
set "LAN_IP=127.0.0.1"
for /f "usebackq delims=" %%I in (`powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ip=(Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.IPAddress -ne '127.0.0.1' -and $_.IPAddress -notlike '169.254*' -and $_.ValidLifetime -gt 0} | Sort-Object InterfaceMetric | Select-Object -First 1 -ExpandProperty IPAddress); if($ip){$ip}else{'127.0.0.1'}"`) do (
  set "LAN_IP=%%I"
)

set "URL_ADMIN=http://%LAN_IP%:%PORT%/login"
set "URL_PLAYER=http://%LAN_IP%:%PORT%/player_join"

REM Giải phóng port 5000 nếu đang bị chiếm
for /f "tokens=5" %%P in ('netstat -ano 2^>nul ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  echo  [INFO] Port %PORT% dang su dung, giai phong...
  taskkill /F /PID %%P >nul 2>&1
)

REM =========================================
REM  Khởi động server Flask ẩn nền
REM =========================================
echo  [1/3] Dang khoi dong server...
del /q server_stdout.log server_stderr.log >nul 2>&1
start "" /B "%PYTHON%" "%APP%" > server_stdout.log 2> server_stderr.log

REM =========================================
REM  Chờ server sẵn sàng (tối đa 40 lần x 0.3s = 12s)
REM =========================================
echo  [2/3] Cho server san sang...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$u='%URL_ADMIN%'; for($i=0;$i-lt40;$i++){try{(Invoke-WebRequest -UseBasicParsing -TimeoutSec 2 $u)|Out-Null;break}catch{};Start-Sleep -Milliseconds 300}" >nul 2>&1

REM =========================================
REM  Mở trình duyệt
REM =========================================
echo  [3/3] Mo trinh duyet...
start "" "%URL_ADMIN%"

REM Lấy PID server để dừng sau
set "SERVER_PID="
for /f "tokens=5" %%P in ('netstat -ano 2^>nul ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  set "SERVER_PID=%%P"
)

echo.
echo  ==========================================
echo   SERVER DANG CHAY (PID: %SERVER_PID%)
echo  ==========================================
echo.
echo   Dia chi Admin (MC):   %URL_ADMIN%
echo   Dia chi nguoi choi:   %URL_PLAYER%
echo.
echo   Cac thiet bi khac ket noi qua Wi-Fi/LAN
echo   va truy cap dia chi nguoi choi ben tren.
echo.
echo   Nhan phim BAT KY de DUNG server va thoat.
echo  ==========================================
pause >nul

REM Dừng server
echo  Dang dung server...
if not "%SERVER_PID%"=="" (
  taskkill /F /PID %SERVER_PID% >nul 2>&1
)
for /f "tokens=5" %%P in ('netstat -ano 2^>nul ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  taskkill /F /PID %%P >nul 2>&1
)

echo  Da dung. Tam biet!
timeout /t 2 >nul
endlocal
exit /b 0
