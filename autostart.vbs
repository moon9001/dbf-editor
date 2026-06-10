Set WshShell = CreateObject("WScript.Shell")
' 静默启动 DBF 编辑器服务器（无命令行窗口）
WshShell.Run """F:\AI\xionglaoshi\0610xiong\dbf_editor\start_dbf_editor.bat""", 0, False
