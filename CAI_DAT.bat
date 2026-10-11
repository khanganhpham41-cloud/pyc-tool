@echo off
chcp 65001 >nul
title CAI DAT CONG CU PHIEU YCPT
cd /d "%~dp0"
echo ==== Kiem tra Python ====
set PY=
where py >nul 2>nul && (py -3 --version >nul 2>nul && set PY=py -3)
if not defined PY ( where python >nul 2>nul && (python --version >nul 2>nul && set PY=python) )
if not defined PY (
  echo Chua co Python. Dang cai Python 3.12 qua winget ^(can mang^)...
  winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
  echo.
  echo Cai xong Python: DONG cua so nay roi bam dup CAI_DAT.bat them 1 lan nua.
  pause
  exit /b
)
echo Dung: %PY%
%PY% --version
echo ==== Cai thu vien python-docx, openpyxl ====
%PY% -m pip install --disable-pip-version-check --no-index --find-links wheels python-docx openpyxl >nul 2>nul
if errorlevel 1 %PY% -m pip install --disable-pip-version-check python-docx openpyxl
%PY% -c "import docx, openpyxl; print('OK - thu vien da san sang')"
if errorlevel 1 ( echo LOI: chua cai duoc thu vien. Chup man hinh gui lai. ) else ( echo. & echo CAI DAT XONG. Tu nay chi can dung TAO_PHIEU.bat )
pause
