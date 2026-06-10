@echo off
chcp 65001 >nul
echo ================================
echo   启用 DBF 编辑器开机自启
echo ================================
echo.
set "vbs=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\DBF编辑器自启.vbs"
copy /y "%~dp0autostart.vbs" "%vbs%" >nul
if %errorlevel%==0 (
    echo [√] 已启用开机自启
    echo     下次开机将自动启动 DBF 编辑器服务器
) else (
    echo [×] 启用失败，请以管理员身份运行
)
echo.
pause
