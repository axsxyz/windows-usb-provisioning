# Provisioning — menu lanceur (fleches). Appele par start.bat apres elevation admin.
# UTF-8 with BOM

$ErrorActionPreference = "Continue"
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8; chcp 65001 | Out-Null } catch {}

# Plein ecran des que la session est admin
try {
    if (-not ("ProvConsoleWin" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class ProvConsoleWin {
  [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  public const int SW_MAXIMIZE = 3;
}
"@
    }
    $hwnd = [ProvConsoleWin]::GetConsoleWindow()
    if ($hwnd -ne [IntPtr]::Zero) { [void][ProvConsoleWin]::ShowWindow($hwnd, 3) }
} catch {}

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptDir = $here
# --- Config : lib\common.ps1 + config.local.psd1 (a cote du script, sinon X:\scripts\ sur la cle) ---
$provLib = @(@(Join-Path $ScriptDir "lib\common.ps1") + @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
    ForEach-Object { Join-Path $_.Root "scripts\lib\common.ps1" }) | Where-Object { Test-Path -LiteralPath $_ }) | Select-Object -First 1
if (-not $provLib) {
    Write-Host "  [ERREUR] lib\common.ps1 introuvable (a cote du script ou X:\scripts\lib\)." -ForegroundColor Red
    Read-Host "  Entree pour quitter" | Out-Null
    exit 1
}
. $provLib
$ProvConfig = Import-ProvConfig -BaseDir (Split-Path -Parent (Split-Path -Parent $provLib))
if (-not $ProvConfig) {
    Read-Host "  Entree pour quitter" | Out-Null
    exit 1
}

$deploy = Join-Path $here "deploy-apps.ps1"
$opts = @(
    [pscustomobject]@{
        Label      = "Deploiement complet du poste"
        File       = $deploy
        Args       = @()
        Hint       = "install apps, metier, Office, rapports..."
        StayInMenu = $false
    }
    [pscustomobject]@{
        Label      = "Modules / retouches (choix)"
        File       = $deploy
        Args       = @('-ModulePicker')
        Hint       = "favoris, demarrage, Chrome defaut, barre, Bureau…"
        StayInMenu = $true
    }
    [pscustomobject]@{
        Label      = "Migration ancien poste (fichiers + mdp Chrome)"
        File       = (Join-Path $here "migrate-user.ps1")
        Args       = @()
        Hint       = "SANS admin — Bureau / Docs / DL + export mots de passe"
        StayInMenu = $false
    }
    [pscustomobject]@{
        Label      = "Rangement rapide du Bureau"
        File       = (Join-Path $here "desktop-organize.ps1")
        Args       = @()
        Hint       = "heros metier + _Inbox / Applications"
        StayInMenu = $true
    }
    [pscustomobject]@{
        Label      = "Nettoyage copies script (Bureau)"
        File       = (Join-Path $here "cleanup-deploy.ps1")
        Args       = @()
        Hint       = "supprime deploy-apps du poste — pas la cle USB"
        StayInMenu = $false
    }
)

$i = 0
function Show-ProvMenu {
    Clear-Host
    Write-Host ""
    Write-Host ("  " + $ProvConfig.OrgName + "  ·  menu lanceur") -ForegroundColor Cyan
    Write-Host "  Haut/Bas  ·  Entree = lancer  ·  Echap = quitter" -ForegroundColor DarkGray
    Write-Host ""
    for ($n = 0; $n -lt $opts.Count; $n++) {
        if ($n -eq $i) {
            Write-Host ("  > " + $opts[$n].Label) -ForegroundColor Black -BackgroundColor Cyan
            Write-Host ("      " + $opts[$n].Hint) -ForegroundColor DarkCyan
        } else {
            Write-Host ("    " + $opts[$n].Label) -ForegroundColor Gray
        }
    }
    Write-Host ""
    Write-Host ("  Dossier : $here") -ForegroundColor DarkGray
    Write-Host ""
}

Show-ProvMenu
while ($true) {
    $k = [Console]::ReadKey($true)
    switch ($k.Key) {
        'UpArrow' {
            $i = ($i - 1 + $opts.Count) % $opts.Count
            Show-ProvMenu
        }
        'DownArrow' {
            $i = ($i + 1) % $opts.Count
            Show-ProvMenu
        }
        'Escape' { exit 0 }
        'Enter' {
            $sel = $opts[$i]
            $f = $sel.File
            if (-not (Test-Path -LiteralPath $f)) {
                Write-Host ""
                Write-Host "  Fichier manquant : $f" -ForegroundColor Red
                Start-Sleep -Seconds 3
                if ($sel.StayInMenu) { Show-ProvMenu; continue }
                exit 1
            }
            Clear-Host
            Write-Host ""
            Write-Host ("  Lancement : " + $sel.Label) -ForegroundColor Cyan
            Write-Host ""
            $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $f) + @($sel.Args)
            & powershell.exe @argList
            if ($sel.StayInMenu) {
                Show-ProvMenu
            } else {
                exit $LASTEXITCODE
            }
        }
    }
}
