@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo ============================================
echo   正在安装打包工具 PyInstaller ...
echo ============================================
"%PY%" -m pip install --disable-pip-version-check pyinstaller

echo.
echo ============================================
echo   正在打包为 WP工作台.exe ...
echo ============================================
"%PY%" -m PyInstaller --noconfirm --clean --onefile ^
  --name "WP工作台" ^
  --distpath "%~dp0." ^
  --workpath "%~dp0build" ^
  --specpath "%~dp0build" ^
  --add-data "%~dp0app\templates;templates" ^
  --add-data "%~dp0app\static;static" ^
  --hidden-import=openpyxl ^
  --hidden-import=lxml ^
  --collect-all openpyxl ^
  "%~dp0app\app.py"

echo.
if exist "WP工作台.exe" (
  echo 打包完成：%~dp0WP工作台.exe
  echo 双击即可运行；data\config.json\backup\exports 会生成在 exe 同目录。
) else (
  echo 打包失败，请查看上面的错误信息。
)
pause
