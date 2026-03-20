@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

cd /d "%~dp0"
title INI-Excel WebView 转换工具 - 打包器

set "VENV_DIR=.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "MAIN_SCRIPT=script\gui.py"
set "APP_NAME=INI-Excel转换工具"
set "SETUP_SCRIPT=setup.bat"

echo ========================================
echo   INI-Excel WebView 转换工具 - 打包
echo ========================================
echo.

if not exist "%SETUP_SCRIPT%" (
    echo [错误] 未找到环境准备脚本：%SETUP_SCRIPT%
    echo.
    pause
    exit /b 1
)

if not exist "%MAIN_SCRIPT%" (
    echo [错误] 未找到主程序：%MAIN_SCRIPT%
    echo.
    pause
    exit /b 1
)

call "%SETUP_SCRIPT%"
if errorlevel 1 (
    echo [错误] setup 执行失败，无法继续打包。
    echo.
    pause
    exit /b 1
)

echo [信息] 安装 PyInstaller ...
"%VENV_PY%" -m pip install pyinstaller
if errorlevel 1 (
    echo [错误] 安装 PyInstaller 失败。
    echo.
    pause
    exit /b 1
)

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "%APP_NAME%.spec" del /f /q "%APP_NAME%.spec"

echo [信息] 开始打包...
"%VENV_PY%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "%APP_NAME%" ^
  --add-data "webui;webui" ^
  --add-data "config;config" ^
  "%MAIN_SCRIPT%"

if errorlevel 1 (
    echo.
    echo [错误] PyInstaller 打包失败。
    echo.
    pause
    exit /b 1
)

echo.
echo [成功] 打包完成！
echo 输出文件位置：
echo     dist\%APP_NAME%.exe
echo.
pause

endlocal
exit /b 0
