#Requires -RunAsAdministrator
# Provisioning — Nettoyage des copies deploy-apps sur le POSTE (jamais la cle USB)
# Usage : start.bat sur la cle → menu "Nettoyage copies script (Bureau)"
# UTF-8 BOM requis pour Windows PowerShell 5.1

$ErrorActionPreference = "Continue"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
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
$UI = @{
    Arrow = [string][char]0x2192
    VLine = [string][char]0x2502
}

function Test-ProvPathOnRemovableDrive {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    try {
        $full = if (Test-Path -LiteralPath $Path) { (Resolve-Path -LiteralPath $Path).Path } else { $Path }
        $root = [IO.Path]::GetPathRoot($full)
        $letter = $root.TrimEnd('\')
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$letter'" -EA SilentlyContinue
        if (-not $disk) { return $false }
        if ($disk.DriveType -eq 2) { return $true }
        foreach ($sub in @("scripts\offline", "offline")) {
            $m = Join-Path $letter $sub
            $probes = @("AnyDesk.exe", $ProvConfig.NoteFileName)
            if ($ProvConfig.PortableApp -and $ProvConfig.PortableApp.Folder) { $probes += [string]$ProvConfig.PortableApp.Folder }
            foreach ($probe in $probes) {
                if (Test-Path (Join-Path $m $probe)) { return $true }
            }
        }
        return $false
    } catch { return $false }
}

$desk = [Environment]::GetFolderPath("Desktop")
$names = @("deploy-apps.ps1", "deploy-apps.bat")
$bases = @($desk, (Join-Path $desk "scripts"))
$targets = @()
foreach ($n in $names) {
    foreach ($b in $bases) {
        $p = Join-Path $b $n
        if ((Test-Path -LiteralPath $p) -and -not (Test-ProvPathOnRemovableDrive $p)) {
            $targets += $p
        }
    }
}

Write-Host ""
Write-Host ("  " + $ProvConfig.OrgName + " — cleanup-deploy (poste uniquement)") -ForegroundColor Cyan
Write-Host "  La cle USB ne sera jamais modifiee." -ForegroundColor DarkGray
Write-Host ""

if ($targets.Count -eq 0) {
    Write-Host "  Rien a supprimer sur le Bureau." -ForegroundColor Green
    # ghost offline
} else {
    Write-Host "  Fichiers trouves :" -ForegroundColor Yellow
    $targets | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
    Write-Host ""
    Write-Host "  Tapez  OUI  pour supprimer + vider la corbeille" -ForegroundColor Yellow
    $ans = (Read-Host "  Confirmation").Trim().ToLowerInvariant()
    if ($ans -in @('o', 'oui', 'y', 'yes')) {
        foreach ($t in $targets) {
            try {
                Remove-Item -LiteralPath $t -Force -ErrorAction Stop
                Write-Host "  OK  $t" -ForegroundColor Green
            } catch {
                Write-Host "  ECHEC  $t — $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        $ds = Join-Path $desk "scripts"
        if ((Test-Path $ds) -and -not (Test-ProvPathOnRemovableDrive $ds)) {
            $left = @(Get-ChildItem $ds -Force -EA SilentlyContinue | Where-Object { $_.Name -notmatch '^\._' })
            if ($left.Count -eq 0) { Remove-Item $ds -Force -Recurse -EA SilentlyContinue }
        }
        try { Clear-RecycleBin -Force -EA Stop; Write-Host "  Corbeille videe" -ForegroundColor Green } catch {}
    } else {
        Write-Host "  Annule." -ForegroundColor DarkGray
    }
}

foreach ($ghost in @((Join-Path $desk "offline"), (Join-Path $desk "scripts\offline"))) {
    if (-not (Test-Path $ghost)) { continue }
    $kids = @(Get-ChildItem $ghost -Force -EA SilentlyContinue | Where-Object { $_.Name -notmatch '^\._' })
    if ($kids.Count -eq 0) {
        Remove-Item $ghost -Force -Recurse -EA SilentlyContinue
        Write-Host "  Ghost offline supprime : $ghost" -ForegroundColor Green
    }
}

$docs = Join-Path ([Environment]::GetFolderPath("MyDocuments")) $ProvConfig.ReportFolderName
if (Test-Path $docs) {
    Write-Host ("  " + $UI.Arrow + " Rapports conserves : $docs") -ForegroundColor Cyan
}
Write-Host ""
Read-Host "  Entree pour fermer"
