@echo off
chcp 65001 >nul
cd /d "%~dp0"
java -Dstdout.encoding=UTF-8 -Dstderr.encoding=UTF-8 -Dfile.encoding=UTF-8 -jar target\spring-app2.jar
echo.
pause
