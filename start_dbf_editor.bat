@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title DBF 编辑器服务器

echo.
echo ================================
echo    DBF 编辑器 - 启动中...
echo ================================
echo.

cd /d "%~dp0"

REM 检查 Python 环境
if not exist "venv\Scripts\python.exe" (
    echo [错误] 未找到 Python 虚拟环境
    echo.
    echo 请先安装依赖，运行以下命令：
    echo.
    echo   python -m venv venv
    echo   venv\Scripts\pip install flask dbf openpyxl
    echo.
    pause
    exit /b 1
)

REM 获取本机局域网 IP
set "LOCAL_IP="
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4" 2^>nul') do (
    set "ip=%%a"
    set "ip=!ip: =!"
    if not "!ip!"=="127.0.0.1" (
        if "!LOCAL_IP!"=="" set "LOCAL_IP=!ip!"
    )
)

echo.
echo   本地访问: http://127.0.0.1:5000
if not "!LOCAL_IP!"=="" echo   局域网访问: http://!LOCAL_IP!:5000
echo.
echo   按 Ctrl+C 停止服务器
echo ================================
echo.

venv\Scripts\python.exe server.py

pause
