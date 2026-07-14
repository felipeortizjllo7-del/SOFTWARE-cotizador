@echo off
cd /d "%~dp0"
where py >nul 2>nul && ( py publicar_version.py %* ) || ( python publicar_version.py %* )
echo.
pause
