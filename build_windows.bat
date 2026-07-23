@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ========================================
echo Excel order downloader - Windows build
echo Power by LingTu (ipoddy.cn)
echo ========================================

where py >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python Launcher not found. Install Python 3.12 x64 first.
  pause
  exit /b 1
)

if not exist ".venv-win\Scripts\python.exe" (
  echo [1/5] Creating Python virtual environment...
  py -3.12 -m venv .venv-win
  if errorlevel 1 goto :failed
)

call ".venv-win\Scripts\activate.bat"

echo [2/5] Installing build dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt
if errorlevel 1 goto :failed

echo [3/5] Installing Windows Chromium...
if exist ".playwright-browsers" rmdir /s /q ".playwright-browsers"
set "PLAYWRIGHT_BROWSERS_PATH=%CD%\.playwright-browsers"
python -m playwright install --only-shell chromium
if errorlevel 1 goto :failed

echo [4/5] Building single-file Windows EXE...
python -m PyInstaller --noconfirm --clean xlsx2files_gui.spec
if errorlevel 1 goto :failed

if not exist "dist\Excel订单素材下载工具.exe" goto :failed

echo [5/5] Preparing customer package...
copy /y "CUSTOMER_GUIDE.txt" "dist\客户使用说明.txt" >nul
powershell -NoProfile -Command "$file='dist\Excel订单素材下载工具.exe'; $hash=(Get-FileHash -Algorithm SHA256 $file).Hash; ($hash + '  Excel订单素材下载工具.exe') | Set-Content -Encoding ascii 'dist\SHA256.txt'"

echo.
echo Build completed:
echo %CD%\dist\Excel订单素材下载工具.exe
echo.
echo Send these three files to the customer:
echo   Excel订单素材下载工具.exe
echo   客户使用说明.txt
echo   SHA256.txt
pause
exit /b 0

:failed
echo.
echo [ERROR] Build failed. Review the messages above.
pause
exit /b 1
