@echo off
chcp 65001 >nul
title TAO PHIEU YEU CAU PHAN TICH
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PY=
where py >nul 2>nul && (py -3 --version >nul 2>nul && set PY=py -3)
if not defined PY ( where python >nul 2>nul && (python --version >nul 2>nul && set PY=python) )
if not defined PY ( echo Chua cai Python - chay CAI_DAT.bat truoc. & pause & exit /b )
%PY% scripts\tao_phieu.py %*
if errorlevel 1 ( echo. & echo Chua tao duoc phieu - xem thong bao o tren. & pause & exit /b )
echo.
echo ==== Word mo va luu lai file (lam sach cau truc, tranh loi "found a problem") ====
powershell -NoProfile -ExecutionPolicy Bypass -File "scripts\lam_sach_bang_word.ps1" "scripts\_danh_sach_file_moi.txt"
if exist "%~dp0HOA_HOC" start "" "%~dp0HOA_HOC"
if exist "%~dp0VI_SINH" start "" "%~dp0VI_SINH"
pause
