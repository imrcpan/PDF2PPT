
@echo off
cd /d "%~dp0PDF2PPT"
PDF2PPT.exe

:: 删除./temp目录及其所有内容
rd /s /q "%~dp0temp"

pause