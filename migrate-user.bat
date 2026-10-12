@echo off
REM ============================================================================
REM  Provisioning — Migration utilisateur (ancien poste)
REM  SANS admin : lancer connecte comme l'utilisateur final
REM ============================================================================
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "%~dp0migrate-user.ps1" (
    echo [ERREUR] migrate-user.ps1 introuvable
    pause
    exit /b 1
)

REM Ne PAS s'auto-elever : les mdp Chrome necessitent le contexte user
net session >nul 2>&1
if not errorlevel 1 (
    echo.
    echo  [ATTENTION] Vous etes en administrateur.
    echo  Fermez et double-cliquez ce .bat SANS "Executer en tant qu'admin",
    echo  connecte avec le compte de l'utilisateur a migrer.
    echo.
    pause
)

chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0migrate-user.ps1"
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" echo  [INFO] Code sortie %RC%.
pause
endlocal & exit /b %RC%
