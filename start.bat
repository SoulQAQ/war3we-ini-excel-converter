@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

cd /d "%~dp0"
title INI-Excel WebView 转换工具 - 启动器

set "VENV_DIR=.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "MAIN_SCRIPT=script\gui.py"
set "REQUIREMENTS=requirements.txt"
set "SNAPSHOT_FILE=%VENV_DIR%\.requirements.snapshot"
set "SETUP_SCRIPT=setup.bat"

echo ========================================
echo   INI-Excel WebView 转换工具
echo ========================================
echo.

if not exist "%MAIN_SCRIPT%" (
    echo [错误] 未找到主程序：%MAIN_SCRIPT%
    echo.
    pause
    exit /b 1
)

set "NEED_SETUP="
if not exist "%VENV_PY%" set "NEED_SETUP=1"
if not exist "%SNAPSHOT_FILE%" set "NEED_SETUP=1"
if exist "%REQUIREMENTS%" (
    fc /b "%REQUIREMENTS%" "%SNAPSHOT_FILE%" >nul 2>nul
    if errorlevel 1 set "NEED_SETUP=1"
)

if defined NEED_SETUP (
    echo [信息] 检测到环境未准备完成或依赖已变化，调用 setup...
    call "%SETUP_SCRIPT%"
    if errorlevel 1 (
        echo [错误] setup 执行失败，程序无法启动。
        echo.
        pause
        exit /b 1
    )
) else (
    echo [信息] 当前环境与 requirements.txt 一致，跳过 setup。
)

echo.
echo [信息] 正在启动程序...
"%VENV_PY%" "%MAIN_SCRIPT%"
set "APP_EXIT=%ERRORLEVEL%"

echo.
if not "%APP_EXIT%"=="0" (
    echo [错误] 程序已退出，退出码：%APP_EXIT%
    echo.
    pause
    exit /b %APP_EXIT%
)

endlocal
exit /b 0
