@echo off
cd /d "%~dp0."
schtasks /Create /TN WP_DDL_DailyCheck /TR "\"%~dp0daily_check.bat\"" /SC DAILY /ST 10:30 /F
echo 已注册任务 WP_DDL_DailyCheck（每天 10:30 运行 daily_check.bat）
pause
