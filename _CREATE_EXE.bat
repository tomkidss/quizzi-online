@echo off
setlocal EnableExtensions
title Quizzi - Dang tao file EXE...
cd /d "%~dp0"

echo.
echo  ========================================
echo    QUIZZI EXE BUILDER (7-Zip SFX)
echo  ========================================
echo.

REM =========================================
REM  Tim 7-Zip (dung if don gian, khong block)
REM =========================================
set "ZIP64=C:\Program Files\7-Zip"
set "ZIP32=C:\Program Files (x86)\7-Zip"

set "SEVENZIP="
set "SFX_MODULE="

if exist "%ZIP64%\7z.exe" set "SEVENZIP=%ZIP64%\7z.exe"
if exist "%ZIP64%\7z.sfx" set "SFX_MODULE=%ZIP64%\7z.sfx"
if exist "%ZIP64%\7zSD.sfx" set "SFX_MODULE=%ZIP64%\7zSD.sfx"
if exist "%ZIP32%\7z.exe" set "SEVENZIP=%ZIP32%\7z.exe"
if exist "%ZIP32%\7z.sfx" set "SFX_MODULE=%ZIP32%\7z.sfx"
if exist "%ZIP32%\7zSD.sfx" set "SFX_MODULE=%ZIP32%\7zSD.sfx"

if "%SEVENZIP%"=="" goto ERR_NO7ZIP
if "%SFX_MODULE%"=="" goto ERR_NOSFX

echo  [OK] 7-Zip  : %SEVENZIP%
echo  [OK] SFX    : %SFX_MODULE%
echo.

REM =========================================
REM  Kiem tra thu muc Portable
REM =========================================
set "PORTABLE=%~dp0Quizzi_V9_PORTABLE"

if not exist "%PORTABLE%\" goto ERR_NOPORTABLE
if not exist "%PORTABLE%\python\python.exe" goto ERR_NOPYTHON

echo  [OK] Portable: %PORTABLE%
echo.

REM =========================================
REM  Cac bien duong dan
REM =========================================
set "OUTPUT=%~dp0Quizzi_V9_Setup.exe"
set "ARCHIVE=%~dp0_qtmp.7z"
set "CONFIG=%~dp0_qsfx.cfg"

REM =========================================
REM  Tao file SFX Config
REM =========================================
echo  [1/4] Tao cau hinh SFX...
if exist "%CONFIG%" del /q "%CONFIG%"

echo ;!@Install@!UTF-8!>> "%CONFIG%"
echo Title="Quizzi Realtime v9.4.8">> "%CONFIG%"
echo BeginPrompt="Cai dat Quizzi Realtime?">> "%CONFIG%"
echo ExtractPath="Quizzi_V9_PORTABLE">> "%CONFIG%"
echo RunProgram="CHAY QUIZZI.bat">> "%CONFIG%"
echo ;!@InstallEnd@!>> "%CONFIG%"

REM =========================================
REM  Nen du lieu bang 7-Zip
REM =========================================
echo  [2/4] Nen du lieu (3-5 phut)...
if exist "%ARCHIVE%" del /q "%ARCHIVE%"

"%SEVENZIP%" a -t7z -mx=5 -mmt=on "%ARCHIVE%" "%PORTABLE%\*" -r
if errorlevel 1 goto ERR_COMPRESS

echo  [OK] Nen xong!
echo.

REM =========================================
REM  Ghep SFX + Config + Archive = EXE
REM =========================================
echo  [3/4] Tao file EXE...
if exist "%OUTPUT%" del /q "%OUTPUT%"

copy /b "%SFX_MODULE%" + "%CONFIG%" + "%ARCHIVE%" "%OUTPUT%" >nul
if errorlevel 1 goto ERR_COPY

REM =========================================
REM  Don dep file tam
REM =========================================
echo  [4/4] Don dep...
if exist "%CONFIG%" del /q "%CONFIG%"
if exist "%ARCHIVE%" del /q "%ARCHIVE%"

for %%F in ("%OUTPUT%") do set "EXE_SIZE=%%~zF"
set /a "EXE_MB=%EXE_SIZE% / 1048576"

echo.
echo  ========================================
echo   HOAN TAT!
echo  ========================================
echo.
echo   File: Quizzi_V9_Setup.exe (~%EXE_MB% MB)
echo.
echo   CACH SU DUNG:
echo   - Gui file "Quizzi_V9_Setup.exe" cho nguoi nhan
echo   - Ho double-click de chay
echo   - Neu SmartScreen canh bao: "More info" ^ "Run anyway"
echo   - App tu giai nen va mo Admin tren trinh duyet
echo   (Khong can cai Python hay gi ca!)
echo  ========================================
echo.
start "" /select,"%OUTPUT%"
pause
endlocal
exit /b 0

REM =========================================
REM  Xu ly loi
REM =========================================
:ERR_NO7ZIP
echo  [LOI] Khong tim thay 7-Zip!
echo  Tai va cai tai: https://www.7-zip.org
pause & exit /b 1

:ERR_NOSFX
echo  [LOI] Khong tim thay module SFX (7z.sfx hoac 7zSD.sfx).
echo  Hay cai lai 7-Zip (full installer, khong phai portable).
pause & exit /b 1

:ERR_NOPORTABLE
echo  [LOI] Khong tim thay thu muc "Quizzi_V9_PORTABLE".
echo  Hay chay "_BUILD_PORTABLE.bat" truoc.
pause & exit /b 1

:ERR_NOPYTHON
echo  [LOI] Thu muc Portable khong hop le (thieu Python).
echo  Hay chay lai "_BUILD_PORTABLE.bat".
pause & exit /b 1

:ERR_COMPRESS
echo  [LOI] Nen du lieu that bai.
if exist "%CONFIG%" del /q "%CONFIG%"
if exist "%ARCHIVE%" del /q "%ARCHIVE%"
pause & exit /b 1

:ERR_COPY
echo  [LOI] Tao file EXE that bai.
if exist "%CONFIG%" del /q "%CONFIG%"
if exist "%ARCHIVE%" del /q "%ARCHIVE%"
pause & exit /b 1
