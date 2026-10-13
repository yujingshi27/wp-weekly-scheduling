@echo off
chcp 65001 >nul
pushd "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo ============================================
echo   WP 工作台 启动中，请稍候...
echo   浏览器会自动打开；关闭本窗口即退出工作台。
echo ============================================
echo.
"%PY%" "%~dp0app\app.py"

echo.
echo 工作台已退出。
popd
pause
