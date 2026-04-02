@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Quizzi - Dang dong goi Portable Bundle...
cd /d "%~dp0"

REM ===================================================
REM   QUIZZI PORTABLE BUILDER
REM   Chay script nay 1 lan tren may dev de tao boi.
REM   Ket qua: thu muc "Quizzi_V9_PORTABLE" san sang gui.
REM ===================================================

set "DIST=.\Quizzi_V9_PORTABLE"
set "PY_VERSION=3.12.9"
set "PY_ZIP=python-%PY_VERSION%-embed-amd64.zip"
set "PY_URL=https://www.python.org/ftp/python/%PY_VERSION%/%PY_ZIP%"
set "PY_DIR=%DIST%\python"
set "GET_PIP_URL=https://bootstrap.pypa.io/get-pip.py"

echo.
echo  ========================================
echo    QUIZZI PORTABLE BUILDER v9.4.8
echo  ========================================
echo.

REM Kiểm tra PowerShell (cần để download)
where powershell >nul 2>&1
if errorlevel 1 (
  echo [LOI] Khong tim thay PowerShell. Can Windows 10+.
  pause & exit /b 1
)

REM =========================================
REM  BUOC 1: Tao cau truc thu muc
REM =========================================
echo [1/6] Tao cau truc thu muc output...
if exist "%DIST%" (
  echo       Thu muc cu da ton tai, xoa va tao lai...
  rmdir /s /q "%DIST%"
)
mkdir "%DIST%"
mkdir "%PY_DIR%"

REM =========================================
REM  BUOC 2: Tai Python Embeddable
REM =========================================
echo [2/6] Tai Python %PY_VERSION% Embeddable...
if not exist "%PY_ZIP%" (
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Write-Host '  Dang tai tu python.org...'; (New-Object Net.WebClient).DownloadFile('%PY_URL%', '%PY_ZIP%')"
  if errorlevel 1 (
    echo [LOI] Tai Python that bai. Kiem tra ket noi internet.
    pause & exit /b 1
  )
) else (
  echo       Da co file %PY_ZIP%, bo qua tai xuong.
)

REM Giải nén Python Embeddable
echo       Giai nen Python vao %PY_DIR%...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Expand-Archive -Path '%PY_ZIP%' -DestinationPath '%PY_DIR%' -Force"

REM =========================================
REM  BUOC 3: Kich hoat pip cho Python Embeddable
REM =========================================
echo [3/6] Kich hoat pip...

REM Sua file .pth de Python Embeddable doc site-packages
set "PTH_FILE=%PY_DIR%\python312._pth"
if not exist "%PTH_FILE%" (
  REM Try other version names
  for %%F in ("%PY_DIR%\python3*._pth") do set "PTH_FILE=%%F"
)

REM Ghi lai file .pth voi import site duoc bat
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$f='%PTH_FILE%'; $c=Get-Content $f; $c=$c -replace '#import site','import site'; Set-Content $f $c"

REM Tai get-pip.py
echo       Tai get-pip.py...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "(New-Object Net.WebClient).DownloadFile('%GET_PIP_URL%', '%PY_DIR%\get-pip.py')"

REM Cai pip
echo       Cai pip vao Python Embeddable...
"%PY_DIR%\python.exe" "%PY_DIR%\get-pip.py" --no-warn-script-location >nul 2>&1

REM =========================================
REM  BUOC 4: Cai tat ca thu vien
REM =========================================
echo [4/6] Cai tat ca thu vien (co the mat 3-5 phut)...
echo       (flask, flask-socketio, openpyxl, qrcode, pillow...)
"%PY_DIR%\python.exe" -m pip install --no-warn-script-location ^
  -r requirements.txt

if errorlevel 1 (
  echo [LOI] Cai thu vien that bai. Kiem tra ket noi internet va requirements.txt
  pause & exit /b 1
)
echo       Cai dat thu vien thanh cong!

REM =========================================
REM  BUOC 5: Copy file ung dung (bo qua venv, cache, db...)
REM =========================================
echo [5/6] Copy file ung dung...

REM Copy file chinh
copy "app.py"          "%DIST%\app.py" >nul
copy "requirements.txt" "%DIST%\requirements.txt" >nul

REM Copy thu muc static va templates
xcopy "static"    "%DIST%\static"    /E /I /Q >nul
xcopy "templates" "%DIST%\templates" /E /I /Q >nul

REM Tao thu muc exports va logs rong
mkdir "%DIST%\exports" >nul 2>&1
mkdir "%DIST%\logs"    >nul 2>&1

REM Copy launcher chinh
copy "CHẠY QUIZZI.bat" "%DIST%\CHẠY QUIZZI.bat" >nul

REM =========================================
REM  BUOC 6: Hoan tat
REM =========================================
echo [6/6] Hoan tat!
echo.

REM Dem kich thuoc
for /f %%S in ('powershell -NoProfile -Command "(Get-ChildItem -Path \"%DIST%\" -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB"') do (
  set "SIZE=%%S"
)

echo  ========================================
echo   BUILD THANH CONG!
echo  ========================================
echo.
echo   Thu muc output: %DIST%
echo   Kich thuoc:     %SIZE% MB (xap xi)
echo.
echo   CAC BUOC TIEP THEO:
echo   1. Nen thu muc "%DIST%" thanh file .zip
echo   2. Gui file .zip den nguoi dung
echo   3. Nguoi dung giai nen va double-click
echo      "CHAY QUIZZI.bat" la chay duoc
echo.
echo   LUU Y: KHONG gui kem thu muc "venv\"
echo          va file "quiz.db" (du lieu rieng tu)
echo  ========================================
echo.

start "" "%DIST%"
pause
endlocal
exit /b 0
