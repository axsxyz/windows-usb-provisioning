@echo off
REM ============================================================================
REM  Provisioning — Lanceur unique
REM  Double-cliquer ici : UAC admin, puis menu fleches (deploy / cleanup)
REM ============================================================================
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "%~dp0menu.ps1" (
    echo [ERREUR] menu.ps1 introuvable
    pause
    exit /b 1
)
if not exist "%~dp0deploy-apps.ps1" (
    echo [ERREUR] deploy-apps.ps1 introuvable
    pause
    exit /b 1
)

net session >nul 2>&1
if errorlevel 1 (
    echo Elevation des droits administrateur...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
        "Start-Process -FilePath '%~f0' -WorkingDirectory '%~dp0' -Verb RunAs"
    exit /b 0
)

chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0menu.ps1"
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" (
    echo  [INFO] Code sortie %RC%.
)
pause
endlocal & exit /b %RC%
