@echo off
chcp 65001 >nul
title CAI TU DONG TAO PHIEU YCPT
cd /d "%~dp0"
set PY=
where py >nul 2>nul && (py -3 --version >nul 2>nul && set PY=py -3)
if not defined PY ( where python >nul 2>nul && (python --version >nul 2>nul && set PY=python) )
if not defined PY ( echo Chua cai Python - chay CAI_DAT.bat truoc. & pause & exit /b )
for /f "delims=" %%i in ('%PY% -c "import sys,os;print(os.path.join(os.path.dirname(sys.executable),'pythonw.exe'))"') do set PYW=%%i
if not exist "%PYW%" ( echo Khong tim thay pythonw.exe & pause & exit /b )
echo ==== Tao loi tat khoi dong cung Windows ====
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Startup')+'\PYC_PT_TU_DONG.lnk'); $s.TargetPath='%PYW%'; $s.Arguments='\"%~dp0scripts\tu_dong.py\"'; $s.WorkingDirectory='%~dp0'; $s.WindowStyle=7; $s.Save()"
echo ==== Bat dau chay ngam ====
start "" "%PYW%" "%~dp0scripts\tu_dong.py"
echo.
echo XONG. Tu nay chi can bo file .json vao HANG_CHO (laptop hoac Google Drive\PYC_PT\HANG_CHO)
echo -> phieu tu tao trong khoang 30 giay, co thong bao goc man hinh.
echo Nhat ky: NHAT_KY_TU_DONG.txt     Tat: TAT_TU_DONG.bat
pause
