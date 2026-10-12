@echo off
REM Compat : ancien nom — redirige vers le menu unique start.bat
cd /d "%~dp0"
if exist "%~dp0start.bat" (
    call "%~dp0start.bat"
    exit /b %ERRORLEVEL%
)
echo [ERREUR] start.bat introuvable
pause
exit /b 1
