@echo off
chcp 65001 >nul
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*tu_dong.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\PYC_PT_TU_DONG.lnk" 2>nul
echo Da tat che do tu dong. Bat lai: CAI_TU_DONG.bat
pause
