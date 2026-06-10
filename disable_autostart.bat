@echo off
chcp 65001 >nul
echo ================================
echo   禁用 DBF 编辑器开机自启
echo ================================
echo.
set "vbs=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\DBF编辑器自启.vbs"
if exist "%vbs%" (
    del /f /q "%vbs%"
    echo [√] 已禁用开机自启
) else (
    echo [!] 开机自启未启用
)
echo.
pause
