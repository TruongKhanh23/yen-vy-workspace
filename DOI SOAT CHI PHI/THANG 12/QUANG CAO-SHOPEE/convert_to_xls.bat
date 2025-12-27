@echo off
REM Gọi PowerShell script để convert Excel từ cùng thư mục với file .bat
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0do_not_touch_convert_excel.ps1"
pause
