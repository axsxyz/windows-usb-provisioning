# ============================================================================
#  Provisioning — Migration utilisateur (ancien poste -> cle USB)
#  Fichier : migrate-user.ps1
#  Usage   : migrate-user.bat  (SANS admin — session de l'utilisateur)
#
#  1) Nom dossier migrationWin11_prenom_nom (suggestion Office/UPN)
#  2) Selection interactive Bureau / Documents / Telechargements (niveau 1)
#  3) Copie arborescence sur la cle + export mots de passe Chrome
#
#  UTF-8 BOM requis (Windows PowerShell 5.1).
# ============================================================================

$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
    chcp 65001 | Out-Null
} catch {}

$UI = @{
    Check = [string][char]0x2714
    Cross = [string][char]0x2716
    Play  = [string][char]0x25B8
    Info  = [string][char]0x2139
    Dot   = [string][char]0x00B7
    HLine = [string][char]0x2500
    TL    = [string][char]0x250C
    BL    = [string][char]0x2514
    VLine = [string][char]0x2502
    Box   = [string][char]0x2610  # ballot box
    BoxOn = [string][char]0x2611  # ballot box with check
}

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

$script:EmailDomain = [string]$ProvConfig.EmailDomain

function Write-MigStatus {
    param([string]$Name, [ValidateSet('OK','FAIL','INFO','RUN','SKIP')][string]$Status, [string]$Detail = "")
    $icon = switch ($Status) {
        'OK'   { $UI.Check }
        'FAIL' { $UI.Cross }
        'RUN'  { $UI.Play }
        'INFO' { $UI.Info }
        'SKIP' { $UI.Dot }
    }
    $col = switch ($Status) {
        'OK'   { 'Green' }
        'FAIL' { 'Red' }
        'RUN'  { 'Cyan' }
        'INFO' { 'DarkGray' }
        'SKIP' { 'Yellow' }
    }
    $line = "  $icon  $Name"
    if ($Detail) { $line += "  —  $Detail" }
    Write-Host $line -ForegroundColor $col
}

function Test-ProvIsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-ProvConsoleMaximized {
    try {
        if (-not ("ProvMigConsoleWin" -as [type])) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class ProvMigConsoleWin {
  [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
        }
        $hwnd = [ProvMigConsoleWin]::GetConsoleWindow()
        if ($hwnd -ne [IntPtr]::Zero) { [void][ProvMigConsoleWin]::ShowWindow($hwnd, 3) }
    } catch {}
}

function ConvertTo-ProvAsciiSlug {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = $Text.Trim().ToLowerInvariant()
    $map = @{
        'à'='a';'á'='a';'â'='a';'ä'='a';'ã'='a';'å'='a'
        'è'='e';'é'='e';'ê'='e';'ë'='e'
        'ì'='i';'í'='i';'î'='i';'ï'='i'
        'ò'='o';'ó'='o';'ô'='o';'ö'='o';'õ'='o'
        'ù'='u';'ú'='u';'û'='u';'ü'='u'
        'ý'='y';'ÿ'='y';'ç'='c';'ñ'='n'
        'œ'='oe';'æ'='ae'
    }
    foreach ($k in $map.Keys) { $t = $t.Replace($k, $map[$k]) }
    $t = $t -replace '[^a-z0-9]+', '_'
    $t = $t.Trim('_')
    return $t
}

function ConvertTo-ProvTitleCase {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = $Text.Trim()
    if ($t.Length -eq 1) { return $t.ToUpperInvariant() }
    return $t.Substring(0,1).ToUpperInvariant() + $t.Substring(1).ToLowerInvariant()
}

function Test-ProvLooksLikePersonName {
    param([string]$Prenom, [string]$Nom)
    if ([string]::IsNullOrWhiteSpace($Prenom) -or [string]::IsNullOrWhiteSpace($Nom)) { return $false }
    if ($Prenom.Trim().Length -lt 2 -or $Nom.Trim().Length -lt 3) { return $false }
    return $true
}

function Get-ProvSuggestedIdentity {
    $sug = [ordered]@{ Prenom = ""; Nom = ""; Email = ""; Source = "" }
    try {
        $upn = (& whoami.exe /upn 2>$null | Out-String).Trim()
        if ($upn -match '^[^@\s]+@[^@\s]+$') {
            $sug.Email = $upn.ToLowerInvariant()
            $sug.Source = "UPN"
        }
    } catch {}

    $officeDisplay = $null
    try {
        foreach ($ip in @(
            "HKCU:\Software\Microsoft\Office\16.0\Common\Identity",
            "HKCU:\Software\Microsoft\Office\15.0\Common\Identity"
        )) {
            if (-not (Test-Path $ip)) { continue }
            $id = Get-ItemProperty -LiteralPath $ip -EA SilentlyContinue
            if ($id) {
                foreach ($prop in @("EmailAddress", "UserName")) {
                    if (($id.PSObject.Properties.Name -contains $prop) -and ("$($id.$prop)" -match '@')) {
                        $sug.Email = ([string]$id.$prop).ToLowerInvariant()
                        if ($sug.Source -ne "UPN") { $sug.Source = "Office" }
                    }
                }
                if (($id.PSObject.Properties.Name -contains "FriendlyName") -and $id.FriendlyName) {
                    $officeDisplay = [string]$id.FriendlyName
                }
            }
            $idDir = Join-Path $ip "Identities"
            if (-not (Test-Path $idDir)) { continue }
            foreach ($sub in @(Get-ChildItem -LiteralPath $idDir -EA SilentlyContinue)) {
                $props = Get-ItemProperty -LiteralPath $sub.PSPath -EA SilentlyContinue
                if (-not $props) { continue }
                foreach ($prop in @("FriendlyName", "DisplayName", "Name")) {
                    if (($props.PSObject.Properties.Name -contains $prop) -and ("$($props.$prop)" -match '\S+\s+\S+')) {
                        $officeDisplay = [string]$props.$prop
                    }
                }
                foreach ($prop in @("EmailAddress", "EmailAddresses", "UserName")) {
                    if (-not ($props.PSObject.Properties.Name -contains $prop)) { continue }
                    if ("$($props.$prop)" -match '([^\s;]+@[^\s;]+)') {
                        $sug.Email = $Matches[1].ToLowerInvariant()
                    }
                }
            }
        }
    } catch {}

    if ($officeDisplay -and $officeDisplay -match '\S+\s+\S+') {
        $parts = $officeDisplay.Trim() -split '\s+', 2
        if (Test-ProvLooksLikePersonName $parts[0] $parts[1]) {
            $sug.Prenom = ConvertTo-ProvTitleCase $parts[0]
            $sug.Nom = ConvertTo-ProvTitleCase $parts[1]
            $sug.Source = "Office"
        }
    }

    if ((-not $sug.Prenom -or -not $sug.Nom) -and $sug.Email -and $sug.Email -match [regex]::Escape("@$($script:EmailDomain)$")) {
        $local = ($sug.Email -split '@')[0]
        if ($local.Length -ge 3) {
            $sug.Nom = ConvertTo-ProvTitleCase ($local.Substring(1))
            if (-not $sug.Source) { $sug.Source = "Email" }
            else { $sug.Source = "$($sug.Source)+Email" }
        }
    }

    if (-not $sug.Prenom -or -not $sug.Nom) {
        try {
            $lu = Get-LocalUser -Name $env:USERNAME -EA SilentlyContinue
            if ($lu -and $lu.FullName -match '\S+\s+\S+') {
                $parts = $lu.FullName.Trim() -split '\s+', 2
                if (Test-ProvLooksLikePersonName $parts[0] $parts[1]) {
                    if (-not $sug.Prenom) { $sug.Prenom = ConvertTo-ProvTitleCase $parts[0] }
                    if (-not $sug.Nom) { $sug.Nom = ConvertTo-ProvTitleCase $parts[1] }
                    if (-not $sug.Source) { $sug.Source = "Profil" }
                }
            }
        } catch {}
    }

    return [pscustomobject]$sug
}

function Get-ProvRemovableRoots {
    $list = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($d in @(Get-CimInstance Win32_LogicalDisk -EA SilentlyContinue)) {
            if (-not $d.DeviceID) { continue }
            # Amovible + parfois USB vu en Fixed
            if ($d.DriveType -in @(2, 3) -and $d.DeviceID -ne $env:SystemDrive) {
                $root = $d.DeviceID.TrimEnd('\') + '\'
                if (Test-Path $root) { [void]$list.Add($root.TrimEnd('\')) }
            }
        }
    } catch {}
    return @($list)
}

function Read-ProvMigrationFolderName {
    $sug = Get-ProvSuggestedIdentity
    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " Identite migration " + ($UI.HLine * 32)) -ForegroundColor Magenta
    if ($sug.Prenom -or $sug.Nom -or $sug.Email) {
        Write-Host ("  " + $UI.VLine + "  Suggestion ($($sug.Source)) : $($sug.Prenom) $($sug.Nom)  $($sug.Email)") -ForegroundColor DarkCyan
    }
    Write-Host ("  " + $UI.VLine + "  Dossier : migrationWin11_<prenom>_<nom>") -ForegroundColor Gray
    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
    Write-Host ""

    $pPrompt = "     Prenom"
    if ($sug.Prenom) { $pPrompt += " [$($sug.Prenom)]" }
    $nPrompt = "     Nom"
    if ($sug.Nom) { $nPrompt += " [$($sug.Nom)]" }
    $pIn = (Read-Host $pPrompt).Trim()
    $nIn = (Read-Host $nPrompt).Trim()
    if ([string]::IsNullOrWhiteSpace($pIn) -and $sug.Prenom) { $pIn = $sug.Prenom }
    if ([string]::IsNullOrWhiteSpace($nIn) -and $sug.Nom) { $nIn = $sug.Nom }

    $pSlug = ConvertTo-ProvAsciiSlug $pIn
    $nSlug = ConvertTo-ProvAsciiSlug $nIn
    if (-not $pSlug) { $pSlug = "prenom" }
    if (-not $nSlug) { $nSlug = "nom" }

    $folderName = "migrationWin11_{0}_{1}" -f $pSlug, $nSlug
    $confirm = (Read-Host "     Nom du dossier [$folderName]").Trim()
    if (-not [string]::IsNullOrWhiteSpace($confirm)) {
        if ($confirm -match '^migrationWin11_') {
            $folderName = ($confirm -replace '[^\w\-]', '_').Trim('_')
        } else {
            $folderName = "migrationWin11_" + (ConvertTo-ProvAsciiSlug $confirm)
        }
    }

    return [pscustomobject]@{
        Prenom     = $pIn
        Nom        = $nIn
        FolderName = $folderName
        Email      = $sug.Email
    }
}

function Read-ProvUsbDestination {
    param([string]$FolderName)
    $usbs = @(Get-ProvRemovableRoots)
    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " Cle USB destination " + ($UI.HLine * 30)) -ForegroundColor Magenta
    if ($usbs.Count -gt 0) {
        $i = 0
        foreach ($u in $usbs) {
            $i++
            Write-Host ("  " + $UI.VLine + "  [$i] $u") -ForegroundColor Cyan
        }
    } else {
        Write-Host ("  " + $UI.VLine + "  Aucune cle detectee — saisir un chemin") -ForegroundColor Yellow
    }
    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
    Write-Host ""

    $default = if ($usbs.Count -ge 1) { Join-Path $usbs[0] $FolderName } else { "" }
    $raw = (Read-Host $(if ($default) { "     Destination [Entree = $default]" } else { "     Destination (ex. E:\$FolderName)" })).Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($raw)) { $raw = $default }
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }

    if ($raw -match '^[A-Za-z]:\\?$') {
        $raw = Join-Path ($raw.TrimEnd('\') + '\') $FolderName
    } elseif ((Test-Path $raw -PathType Container) -and ((Split-Path $raw -Leaf) -ne $FolderName)) {
        # Racine USB choisie -> sous-dossier user
        if ($raw.Length -le 3) {
            $raw = Join-Path $raw $FolderName
        }
    }

    try {
        New-Item -ItemType Directory -Path $raw -Force -ErrorAction Stop | Out-Null
        $probe = Join-Path $raw (".{0}-write-test" -f $ProvConfig.FilePrefix)
        Set-Content $probe "ok" -EA Stop
        Remove-Item $probe -Force -EA SilentlyContinue
        Write-MigStatus -Name "Destination" -Status OK -Detail $raw
        return $raw
    } catch {
        Write-MigStatus -Name "Destination" -Status FAIL -Detail $_.Exception.Message
        return $null
    }
}

function Test-ProvIsAppLikeItem {
    param([System.IO.FileSystemInfo]$Item)
    if ($Item.PSIsContainer) { return $false }
    $ext = $Item.Extension.ToLowerInvariant()
    if ($ext -in @('.exe', '.msi', '.msix', '.appx', '.scr', '.com', '.msp')) { return $true }
    if ($ext -eq '.lnk') {
        try {
            $sh = New-Object -ComObject WScript.Shell
            $target = [string]$sh.CreateShortcut($Item.FullName).TargetPath
            if ($target -match '(?i)\\Program Files( \(x86\))?\\|\\WindowsApps\\|\\AppData\\Local\\Programs\\|\\Windows\\System32\\') {
                return $true
            }
            if ($target -match '(?i)\.(exe|msi)$') { return $true }
        } catch { return $true }
    }
    return $false
}

function Get-ProvTopLevelItems {
    param([string]$Path, [switch]$SkipApps)
    $list = New-Object System.Collections.Generic.List[psobject]
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return @()
    }
    foreach ($item in @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)) {
        if ($item.Name -match '^(desktop\.ini|Thumbs\.db|\.DS_Store|~\$)') { continue }
        if ($SkipApps -and (Test-ProvIsAppLikeItem $item)) { continue }
        $size = [int64]0
        try {
            if ($item.PSIsContainer) {
                $sum = (Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Force -ErrorAction SilentlyContinue |
                    Measure-Object -Property Length -Sum).Sum
                if ($null -ne $sum) { $size = [int64]$sum }
            } else {
                $size = [int64]$item.Length
            }
        } catch { $size = [int64]0 }
        [void]$list.Add([pscustomobject]@{
            Name     = [string]$item.Name
            FullName = [string]$item.FullName
            IsDir    = [bool]$item.PSIsContainer
            Size     = $size
            Selected = $true
        })
    }
    # ToArray() : pas de return $out (sinon 1 seul element = tout le tableau)
    if ($list.Count -eq 0) { return @() }
    return $list.ToArray()
}

function Format-ProvSize {
    param($Bytes)
    $n = [int64]0
    try {
        if ($null -eq $Bytes) { $n = 0 }
        elseif ($Bytes -is [Array]) {
            $n = [int64](@($Bytes) | Measure-Object -Sum).Sum
        } else {
            $n = [int64]$Bytes
        }
    } catch { $n = 0 }
    if ($n -ge 1GB) { return ("{0:N1} Go" -f ($n / 1GB)) }
    if ($n -ge 1MB) { return ("{0:N1} Mo" -f ($n / 1MB)) }
    if ($n -ge 1KB) { return ("{0:N0} Ko" -f ($n / 1KB)) }
    return "$n o"
}

function Show-ProvCheckList {
    <#
    Navigation fleches, Espace = cocher, A = tout, D = rien, Entree = valider, Echap = annuler section.
    Retourne les items Selected=$true (ou @() si skip).
    #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Items
    )
    if ($null -eq $Items) { $Items = @() }
    # Aplatir si un appelant a emballe le tableau en 1 element
    $flat = New-Object System.Collections.Generic.List[psobject]
    foreach ($it in @($Items)) {
        if ($null -eq $it) { continue }
        if (($it -is [System.Array]) -and -not ($it.PSObject.Properties.Name -contains 'Name')) {
            foreach ($sub in @($it)) {
                if ($null -ne $sub -and ($sub.PSObject.Properties.Name -contains 'Name')) {
                    [void]$flat.Add($sub)
                }
            }
        } elseif ($it.PSObject.Properties.Name -contains 'Name') {
            [void]$flat.Add($it)
        }
    }
    if ($flat.Count -eq 0) {
        Write-MigStatus -Name $Title -Status SKIP -Detail "aucun element"
        return @()
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($it in $flat) {
        $sz = [int64]0
        try {
            if ($it.Size -is [Array]) { $sz = [int64](@($it.Size) | Measure-Object -Sum).Sum }
            else { $sz = [int64]$it.Size }
        } catch { $sz = 0 }
        [void]$rows.Add([pscustomobject]@{
            Name     = [string]$it.Name
            FullName = [string]$it.FullName
            IsDir    = [bool]$it.IsDir
            Size     = $sz
            Selected = [bool]$it.Selected
        })
    }
    $idx = 0
    $scroll = 0
    $page = 18

    while ($true) {
        Clear-Host
        Write-Host ""
        Write-Host ("  " + $Title) -ForegroundColor Cyan
        Write-Host ("  Haut/Bas  ·  Espace = cocher  ·  A = tout  ·  D = rien  ·  Entree = OK  ·  Echap = skip") -ForegroundColor DarkGray
        Write-Host ""
        $selCount = @($rows | Where-Object { $_.Selected }).Count
        Write-Host ("  Selection : $selCount / $($rows.Count)") -ForegroundColor Yellow
        Write-Host ""

        if ($idx -lt $scroll) { $scroll = $idx }
        if ($idx -ge ($scroll + $page)) { $scroll = $idx - $page + 1 }
        $end = [Math]::Min($rows.Count - 1, $scroll + $page - 1)

        for ($n = $scroll; $n -le $end; $n++) {
            $r = $rows[$n]
            $mark = if ($r.Selected) { "[X]" } else { "[ ]" }
            $kind = if ($r.IsDir) { "DIR " } else { "FILE" }
            $line = "  $mark  $kind  $($r.Name)  ($((Format-ProvSize $r.Size)))"
            if ($n -eq $idx) {
                Write-Host $line -ForegroundColor Black -BackgroundColor Cyan
            } else {
                Write-Host $line -ForegroundColor $(if ($r.Selected) { 'Green' } else { 'Gray' })
            }
        }
        if ($rows.Count -gt $page) {
            Write-Host ("  … $($scroll+1)-$($end+1) / $($rows.Count)") -ForegroundColor DarkGray
        }
        Write-Host ""

        $k = [Console]::ReadKey($true)
        switch ($k.Key) {
            'UpArrow'   { if ($idx -gt 0) { $idx-- } }
            'DownArrow' { if ($idx -lt ($rows.Count - 1)) { $idx++ } }
            'Spacebar'  { $rows[$idx].Selected = -not $rows[$idx].Selected }
            'A'         { foreach ($r in $rows) { $r.Selected = $true } }
            'D'         { foreach ($r in $rows) { $r.Selected = $false } }
            'Escape'    { Clear-Host; return @() }
            'Enter'     { Clear-Host; return @($rows | Where-Object { $_.Selected }) }
        }
    }
}

function Get-ProvRobocopyExitMeaning {
    param([int]$Code)
    $parts = New-Object System.Collections.Generic.List[string]
    if ($Code -eq 0) { return "rien a copier (src vide ou deja a jour)" }
    if ($Code -band 1) { [void]$parts.Add("fichiers copies") }
    if ($Code -band 2) { [void]$parts.Add("extras") }
    if ($Code -band 4) { [void]$parts.Add("mismatch") }
    if ($Code -band 8) { [void]$parts.Add("ECHEC fichiers (verrouilles / chemins longs / >4Go FAT32 / droits)") }
    if ($Code -band 16) { [void]$parts.Add("erreur grave (USB plein / chemin invalide)") }
    if ($parts.Count -eq 0) { return "code $Code" }
    return ($parts -join " + ")
}

function Get-ProvDriveInfo {
    param([string]$AnyPath)
    try {
        $full = $AnyPath
        try {
            $parent = if (Test-Path -LiteralPath $AnyPath) { $AnyPath } else { Split-Path $AnyPath -Parent }
            if ($parent -and (Test-Path -LiteralPath $parent)) {
                $full = (Resolve-Path -LiteralPath $parent).Path
            }
        } catch {}
        $root = [IO.Path]::GetPathRoot($full)
        if ($root -match '^([A-Za-z]):') {
            $letter = $Matches[1]
            $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$letter`:'" -EA SilentlyContinue
            if ($disk) {
                return [pscustomobject]@{
                    Letter     = $letter
                    Free       = [int64]$disk.FreeSpace
                    Size       = [int64]$disk.Size
                    FileSystem = [string]$disk.FileSystem
                }
            }
        }
    } catch {}
    return $null
}

function Test-ProvDestinationFreeSpace {
    param([string]$DestPath, [int64]$NeedBytes)
    $disk = Get-ProvDriveInfo -AnyPath $DestPath
    if (-not $disk -or $disk.Free -lt 0) {
        return [pscustomobject]@{ Ok = $true; Free = -1; Need = $NeedBytes; FileSystem = $null }
    }
    if ($disk.Free -lt $NeedBytes) {
        return [pscustomobject]@{ Ok = $false; Free = $disk.Free; Need = $NeedBytes; FileSystem = $disk.FileSystem }
    }
    return [pscustomobject]@{ Ok = $true; Free = $disk.Free; Need = $NeedBytes; FileSystem = $disk.FileSystem }
}

function Find-ProvFilesOverLimit {
    <# Fichiers > limite (ex. 4Go FAT32) sous un dossier. #>
    param([string]$Path, [int64]$LimitBytes = 4GB)
    $hits = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    try {
        Get-ChildItem -LiteralPath $Path -Recurse -File -Force -EA SilentlyContinue |
            Where-Object { $_.Length -gt $LimitBytes } |
            Select-Object -First 15 |
            ForEach-Object {
                [void]$hits.Add([pscustomobject]@{ Name = $_.FullName; Size = [int64]$_.Length })
            }
    } catch {}
    return ,$hits.ToArray()
}

function Invoke-ProvRobocopy {
    <#
      Appel robocopy FIABLE (pas Start-Process) — chemins avec espaces OK.
      Reprend une copie partielle (ne wipe pas la destination).
      Gros volumes : /J unbuffered, retries, MT modere (USB).
    #>
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination,
        [string]$FilePattern = $null,  # fichier unique : nom seul + Source = dossier parent
        [switch]$ShowProgress
    )
    if (-not (Test-Path -LiteralPath $Source)) {
        return [pscustomobject]@{ Code = 16; Ok = $false; Detail = "source introuvable" }
    }
    try {
        $destParent = Split-Path $Destination -Parent
        if ($destParent -and -not (Test-Path -LiteralPath $destParent)) {
            New-Item -ItemType Directory -Path $destParent -Force | Out-Null
        }
        if (-not $FilePattern -and -not (Test-Path -LiteralPath $Destination)) {
            New-Item -ItemType Directory -Path $Destination -Force | Out-Null
        }
    } catch {}

    # Arguments en tableau natif — & robocopy quote correctement
    $roboArgs = New-Object System.Collections.Generic.List[string]
    [void]$roboArgs.Add($Source)
    [void]$roboArgs.Add($Destination)
    if ($FilePattern) {
        [void]$roboArgs.Add($FilePattern)
    } else {
        [void]$roboArgs.Add("/E")          # sous-dossiers + vides
    }
    [void]$roboArgs.Add("/COPY:DAT")       # data+attr+time (pas ACL — plus simple vers USB)
    [void]$roboArgs.Add("/FFT")            # timestamps FAT/exFAT
    [void]$roboArgs.Add("/XJ")             # ignore junctions (boucles OneDrive)
    [void]$roboArgs.Add("/R:5")            # retries (fichiers verrouilles)
    [void]$roboArgs.Add("/W:3")
    [void]$roboArgs.Add("/J")              # I/O non bufferise = mieux pour gros fichiers
    [void]$roboArgs.Add("/MT:4")           # parallelisme doux (USB)
    if ($ShowProgress) {
        [void]$roboArgs.Add("/ETA")
        [void]$roboArgs.Add("/BYTES")
    } else {
        [void]$roboArgs.Add("/NFL")
        [void]$roboArgs.Add("/NDL")
        [void]$roboArgs.Add("/NJH")
        [void]$roboArgs.Add("/NJS")
        [void]$roboArgs.Add("/NP")
    }

    $logFail = Join-Path $env:TEMP ("{0}-robocopy-{1}.log" -f $ProvConfig.FilePrefix, [guid]::NewGuid().ToString('N').Substring(0,8))
    [void]$roboArgs.Add("/TEE")
    [void]$roboArgs.Add("/LOG+:$logFail")

    Write-Host ("  " + $UI.VLine + "      robocopy → $Destination") -ForegroundColor DarkGray
    Write-Host ("  " + $UI.VLine + "      (gros volume : laisser tourner, reprise auto si relance)") -ForegroundColor DarkGray

    $code = 0
    try {
        # Critique : appel direct, PAS Start-Process (casse les chemins avec espaces)
        & robocopy.exe @($roboArgs.ToArray())
        $code = [int]$LASTEXITCODE
    } catch {
        return [pscustomobject]@{ Code = 16; Ok = $false; Detail = $_.Exception.Message; Log = $logFail }
    }

    # Robocopy : 0-7 = succes (bits info) ; >=8 = echecs
    $ok = ($code -le 7)
    return [pscustomobject]@{
        Code   = $code
        Ok     = $ok
        Detail = (Get-ProvRobocopyExitMeaning -Code $code)
        Log    = $logFail
    }
}

function Copy-ProvSelectedItems {
    param(
        [string]$Label,
        [string]$DestSubDir,
        [object[]]$Items
    )
    if (-not $Items -or $Items.Count -eq 0) {
        Write-MigStatus -Name $Label -Status SKIP -Detail "rien a copier"
        return 0
    }
    New-Item -ItemType Directory -Path $DestSubDir -Force | Out-Null

    $need = [int64]0
    foreach ($it in @($Items)) {
        try { $need += [int64]$it.Size } catch {}
    }
    $disk = Get-ProvDriveInfo -AnyPath $DestSubDir
    $fs = if ($disk) { $disk.FileSystem } else { "?" }
    if ($need -gt 0) {
        $space = Test-ProvDestinationFreeSpace -DestPath $DestSubDir -NeedBytes ($need + 64MB)
        if (-not $space.Ok) {
            Write-MigStatus -Name $Label -Status FAIL -Detail ("USB plein ? libre=$(Format-ProvSize $space.Free) besoin~$(Format-ProvSize $space.Need)")
            Write-Host ("  " + $UI.VLine + "  Liberer de l'espace ou utiliser une autre cle, puis relancer.") -ForegroundColor Yellow
            return 0
        }
        Write-Host ("  " + $UI.VLine + "  FS=$fs · ~$(Format-ProvSize $need) a copier · $(Format-ProvSize $space.Free) libres") -ForegroundColor DarkGray
        if ($fs -match '(?i)^FAT32$') {
            Write-Host ("  " + $UI.VLine + "  ATTENTION cle FAT32 : fichiers > 4 Go impossibles — preferer NTFS/exFAT") -ForegroundColor Yellow
        }
    }

    $ok = 0
    $i = 0
    $total = @($Items).Count
    foreach ($it in $Items) {
        $i++
        $target = Join-Path $DestSubDir $it.Name
        $tag = if ($it.IsDir) { "DIR" } else { "FILE" }
        $showProg = [bool]($it.IsDir -and ([int64]$it.Size -gt 200MB))
        Write-Host ("  " + $UI.VLine + "  [$i/$total] $tag  $($it.Name)  ($(Format-ProvSize $it.Size))") -ForegroundColor Cyan
        try {
            if ($it.IsDir) {
                if (-not (Test-Path -LiteralPath $it.FullName)) {
                    Write-MigStatus -Name $it.Name -Status FAIL -Detail "source introuvable"
                    continue
                }

                # Preflight FAT32 + gros fichiers
                if ($fs -match '(?i)^FAT32$') {
                    $over = @(Find-ProvFilesOverLimit -Path $it.FullName -LimitBytes 4GB)
                    if ($over.Count -gt 0) {
                        Write-MigStatus -Name $it.Name -Status FAIL -Detail ("FAT32 : $($over.Count)+ fichier(s) > 4 Go")
                        foreach ($o in @($over | Select-Object -First 5)) {
                            Write-Host ("  " + $UI.VLine + "      > $(Format-ProvSize $o.Size)  $($o.Name)") -ForegroundColor Yellow
                        }
                        Write-Host ("  " + $UI.VLine + "      Reformater la cle en exFAT/NTFS (sauvegarde d'abord) ou copier ces fichiers ailleurs") -ForegroundColor Yellow
                        continue
                    }
                }

                # PAS de wipe destination : robocopy reprend / complete (comme une copie manuelle)
                $result = Invoke-ProvRobocopy -Source $it.FullName -Destination $target -ShowProgress:$showProg
                if ($result.Ok) {
                    $ok++
                    Write-MigStatus -Name $it.Name -Status OK -Detail ("robocopy $($result.Code) — $($result.Detail)")
                } else {
                    Write-MigStatus -Name $it.Name -Status FAIL -Detail ("robocopy $($result.Code) — $($result.Detail)")
                    Write-Host ("  " + $UI.VLine + "      Copie partielle gardee — fermer apps/OneDrive puis RE-lancer (reprise)") -ForegroundColor Yellow
                    if ($result.Log -and (Test-Path -LiteralPath $result.Log)) {
                        Write-Host ("  " + $UI.VLine + "      Log : $($result.Log)") -ForegroundColor DarkGray
                        try {
                            Get-Content -LiteralPath $result.Log -Tail 10 -EA SilentlyContinue | ForEach-Object {
                                if ($_ -and $_.Trim()) {
                                    Write-Host ("  " + $UI.VLine + "        " + $_.Trim()) -ForegroundColor DarkGray
                                }
                            }
                        } catch {}
                    }
                }
            } else {
                # Gros fichier unique : aussi via robocopy (Copy-Item fragile >~2-4 Go)
                $big = ([int64]$it.Size -gt 512MB)
                if ($fs -match '(?i)^FAT32$' -and [int64]$it.Size -gt 4GB) {
                    Write-MigStatus -Name $it.Name -Status FAIL -Detail "FAT32 : fichier > 4 Go"
                    continue
                }
                if ($big) {
                    $parent = Split-Path $it.FullName -Parent
                    $result = Invoke-ProvRobocopy -Source $parent -Destination $DestSubDir -FilePattern $it.Name -ShowProgress
                    if ($result.Ok) {
                        $ok++
                        Write-MigStatus -Name $it.Name -Status OK -Detail ("robocopy $($result.Code)")
                    } else {
                        Write-MigStatus -Name $it.Name -Status FAIL -Detail ("robocopy $($result.Code) — $($result.Detail)")
                    }
                } else {
                    Copy-Item -LiteralPath $it.FullName -Destination $target -Force -EA Stop
                    $ok++
                }
            }
        } catch {
            Write-MigStatus -Name $it.Name -Status FAIL -Detail $_.Exception.Message
        }
    }
    Write-MigStatus -Name $Label -Status OK -Detail "$ok / $total vers $DestSubDir"
    return $ok
}

function Get-ProvChromeProfiles {
    <# Liste Default / Profile N + nom affiche (Local State / Preferences). #>
    $userData = Join-Path $env:LOCALAPPDATA "Google\Chrome\User Data"
    if (-not (Test-Path -LiteralPath $userData)) { return @() }

    $names = @{}
    $localState = Join-Path $userData "Local State"
    if (Test-Path -LiteralPath $localState) {
        try {
            $ls = Get-Content -LiteralPath $localState -Raw -Encoding UTF8 | ConvertFrom-Json
            $cache = $ls.profile.info_cache
            if ($cache) {
                foreach ($prop in $cache.PSObject.Properties) {
                    $n = [string]$prop.Value.name
                    if ($n) { $names[$prop.Name] = $n }
                }
            }
        } catch {
            Write-Host ("  " + $UI.VLine + "  Local State Chrome illisible — noms = dossiers") -ForegroundColor DarkYellow
        }
    }

    $out = New-Object System.Collections.Generic.List[object]
    $dirs = @(Get-ChildItem -LiteralPath $userData -Directory -EA SilentlyContinue |
        Where-Object { $_.Name -eq 'Default' -or $_.Name -match '^Profile \d+$' } |
        Sort-Object { if ($_.Name -eq 'Default') { 0 } else { 1 } }, Name)
    foreach ($d in $dirs) {
        $label = $null
        if ($names.ContainsKey($d.Name)) { $label = $names[$d.Name] }
        if (-not $label) {
            # Fallback Preferences du profil
            try {
                $prefPath = Join-Path $d.FullName "Preferences"
                if (Test-Path -LiteralPath $prefPath) {
                    $pref = Get-Content -LiteralPath $prefPath -Raw -Encoding UTF8 | ConvertFrom-Json
                    if ($pref.profile.name) { $label = [string]$pref.profile.name }
                    elseif ($pref.account_info -and $pref.account_info.Count -gt 0 -and $pref.account_info[0].email) {
                        $label = [string]$pref.account_info[0].email
                    }
                }
            } catch {}
        }
        if (-not $label) { $label = $d.Name }
        [void]$out.Add([pscustomobject]@{
            Directory   = $d.Name
            Path        = $d.FullName
            DisplayName = $label
            FileSuffix  = $(if ($d.Name -eq 'Default') { "" } else { "-" + ($d.Name -replace '\s', '_') })
        })
    }
    return ,@($out.ToArray())
}

function Start-ProvChromeProfile {
    <# Lance Chrome sur un profil precis (ProcessStartInfo = quotes OK pour "Profile 1"). #>
    param(
        [Parameter(Mandatory)][string]$ChromeExe,
        [Parameter(Mandatory)][string]$ProfileDirectory,
        [string]$Url = ""
    )
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ChromeExe
        $args = '--profile-directory="{0}"' -f $ProfileDirectory
        if ($Url) { $args = "$args $Url" }
        $psi.Arguments = $args
        $psi.UseShellExecute = $true
        [void][System.Diagnostics.Process]::Start($psi)
        return $true
    } catch {
        try {
            Start-Process -FilePath $ChromeExe -ArgumentList @("--profile-directory=$ProfileDirectory", $Url)
            return $true
        } catch {
            return $false
        }
    }
}

function Find-ProvRecentChromePasswordCsv {
    param([string]$DestDir, [datetime]$NotBefore)
    $dirs = @($DestDir, (Join-Path $env:USERPROFILE "Downloads"), [Environment]::GetFolderPath("Desktop"))
    $cands = @()
    foreach ($d in $dirs) {
        if (-not (Test-Path $d)) { continue }
        Get-ChildItem $d -Filter "*.csv" -File -EA SilentlyContinue |
            Where-Object {
                $_.LastWriteTime -gt $NotBefore -and
                $_.Name -notmatch ('(?i)^' + [regex]::Escape($ProvConfig.FilePrefix) + '-') -and
                $_.Name -notmatch '(?i)^chrome-passwords'
            } |
            ForEach-Object { $cands += $_ }
    }
    return @($cands | Sort-Object LastWriteTime -Descending)
}

function Export-ProvChromePasswords {
    <#
    Export officiel Chrome (ABE) — un passage par profil si plusieurs.
    Sortie : chrome-passwords.csv (Default) + chrome-passwords-Profile_1.csv …
    #>
    param([string]$DestDir)

    $chrome = @(
        "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $chrome) {
        Write-MigStatus -Name "Chrome" -Status SKIP -Detail "non installe — pas d'export mdp"
        return $false
    }

    $profiles = @(Get-ProvChromeProfiles)
    if ($profiles.Count -eq 0) {
        Write-MigStatus -Name "Mots de passe" -Status SKIP -Detail "aucun profil Chrome"
        return $false
    }

    Get-Process chrome -EA SilentlyContinue | ForEach-Object {
        try { $_.CloseMainWindow() | Out-Null } catch {}
    }
    Start-Sleep 2
    Get-Process chrome -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue

    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " Mots de passe Chrome " + ($UI.HLine * 28)) -ForegroundColor Magenta
    Write-Host ("  " + $UI.VLine + "  Profils detectes : $($profiles.Count)") -ForegroundColor Gray
    foreach ($pr in $profiles) {
        Write-Host ("  " + $UI.VLine + "    - $($pr.DisplayName)  [$($pr.Directory)]") -ForegroundColor DarkCyan
    }
    Write-Host ("  " + $UI.VLine + "  Pour CHAQUE profil choisi : Exporter → mdp WINDOWS → CSV") -ForegroundColor Yellow
    Write-Host ("  " + $UI.VLine + "  puis O ici pour recuperer le fichier.") -ForegroundColor Yellow
    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
    Write-Host ""

    $selected = @()
    if ($profiles.Count -eq 1) {
        $selected = @($profiles[0])
    } else {
        Write-Host "     Quels profils exporter ?" -ForegroundColor Cyan
        Write-Host "       [A] Tous   [Entree = tous]" -ForegroundColor DarkGray
        $i = 0
        foreach ($pr in $profiles) {
            $i++
            Write-Host ("       [$i] $($pr.DisplayName) ($($pr.Directory))") -ForegroundColor Gray
        }
        Write-Host "       Numeros separes par virgule ex. 1,3" -ForegroundColor DarkGray
        $pick = (Read-Host "     Choix").Trim()
        if ([string]::IsNullOrWhiteSpace($pick) -or $pick -match '^(?i)a|all|tous$') {
            $selected = @($profiles)
        } else {
            foreach ($part in ($pick -split '[,;\s]+')) {
                $n = 0
                if ([int]::TryParse($part, [ref]$n) -and $n -ge 1 -and $n -le $profiles.Count) {
                    $selected += $profiles[$n - 1]
                }
            }
            if ($selected.Count -eq 0) { $selected = @($profiles) }
        }
    }

    $anyOk = $false
    foreach ($pr in $selected) {
        $outCsv = Join-Path $DestDir ("chrome-passwords{0}.csv" -f $pr.FileSuffix)
        # Alias stable Default
        if ($pr.Directory -eq 'Default') {
            $outCsv = Join-Path $DestDir "chrome-passwords.csv"
        }

        Write-Host ""
        Write-Host ("  --- Profil : $($pr.DisplayName) [$($pr.Directory)] ---") -ForegroundColor Cyan
        $do = (Read-Host "     Exporter les mdp de ce profil ? [O/n/S=skip rest]").Trim().ToLowerInvariant()
        if ($do -in @('s', 'skip')) { break }
        if ($do -in @('n', 'non', 'no')) {
            Write-MigStatus -Name "Mdp ($($pr.DisplayName))" -Status SKIP -Detail "ignore"
            continue
        }

        Get-Process chrome -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
        Start-Sleep 1
        $before = (Get-Date).AddMinutes(-1)
        try {
            # Ouvre CE profil precis (sinon Chrome reste sur le dernier utilise)
            [void](Start-ProvChromeProfile -ChromeExe $chrome -ProfileDirectory $pr.Directory -Url "chrome://password-manager/settings")
        } catch {}

        Write-Host ("     1) Exporter les mots de passe  2) mdp WINDOWS  3) sauver CSV") -ForegroundColor Yellow
        Write-Host ("     Cible cle : $outCsv") -ForegroundColor DarkCyan

        $found = $null
        while (-not $found) {
            $ans = (Read-Host "     Export [$($pr.DisplayName)] termine ? [O = CSV / N = reouvrir / S = skip profil]").Trim().ToLowerInvariant()
            if ($ans -in @('s', 'skip')) {
                Write-MigStatus -Name "Mdp ($($pr.DisplayName))" -Status SKIP -Detail "ignore"
                break
            }
            if ($ans -in @('n', 'non')) {
                try {
                    [void](Start-ProvChromeProfile -ChromeExe $chrome -ProfileDirectory $pr.Directory -Url "chrome://password-manager/settings")
                } catch {}
                continue
            }
            $cands = @(Find-ProvRecentChromePasswordCsv -DestDir $DestDir -NotBefore $before)
            if ($cands.Count -eq 0) {
                Write-MigStatus -Name "CSV" -Status FAIL -Detail "introuvable — reessayez"
                continue
            }
            if ($cands.Count -eq 1) {
                $found = $cands[0]
            } else {
                Write-Host "     CSV candidats :" -ForegroundColor Cyan
                $j = 0
                foreach ($c in $cands) {
                    $j++
                    Write-Host ("       [$j] $($c.Name)  $($c.LastWriteTime)") -ForegroundColor Gray
                }
                $pick = (Read-Host "     Numero [1]").Trim()
                if ([string]::IsNullOrWhiteSpace($pick)) { $pick = "1" }
                $n = 0; [void][int]::TryParse($pick, [ref]$n)
                if ($n -lt 1 -or $n -gt $cands.Count) { $n = 1 }
                $found = $cands[$n - 1]
            }
        }

        if (-not $found) { continue }

        Copy-Item $found.FullName $outCsv -Force
        Write-MigStatus -Name "Mdp ($($pr.DisplayName))" -Status OK -Detail ("→ " + (Split-Path $outCsv -Leaf))
        $anyOk = $true

        try {
            $srcFull = [IO.Path]::GetFullPath($found.FullName)
            $dstFull = [IO.Path]::GetFullPath($outCsv)
            if ($srcFull -ne $dstFull -and (Test-Path -LiteralPath $found.FullName)) {
                Remove-Item -LiteralPath $found.FullName -Force -ErrorAction Stop
                Write-MigStatus -Name "CSV original" -Status OK -Detail ("supprime : " + $found.Name)
            }
        } catch {
            Write-MigStatus -Name "CSV original" -Status FAIL -Detail $_.Exception.Message
        }

        Get-Process chrome -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
        Start-Sleep 1
    }

    return $anyOk
}

function ConvertTo-ProvChromeBookmarksHtml {
    <#
      JSON Bookmarks Chrome → HTML Netscape.
      Important : la barre (bookmark_bar) est emise avec PERSONAL_TOOLBAR_FOLDER=true
      et SES ENFANTS directement — sinon Chrome cree un dossier "Barre de favoris".
    #>
    param([string]$BookmarksJsonPath, [string]$OutHtmlPath)
    if (-not (Test-Path -LiteralPath $BookmarksJsonPath)) { return $false }
    try {
        $data = Get-Content -LiteralPath $BookmarksJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch { return $false }

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('<!DOCTYPE NETSCAPE-Bookmark-file-1>')
    [void]$lines.Add('<!-- This is an automatically generated file. -->')
    [void]$lines.Add('<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">')
    [void]$lines.Add('<TITLE>Bookmarks</TITLE>')
    [void]$lines.Add('<H1>Bookmarks</H1>')
    [void]$lines.Add('<DL><p>')

    function Prov-Esc([string]$s) {
        if ($null -eq $s) { return "" }
        return ($s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
    }

    $script:ProvBmLines = $lines
    $writer = {
        param($node, $depth)
        if ($null -eq $node) { return }
        $type = [string]$node.type
        $pad = "    " * [Math]::Max(0, $depth)
        if ($type -eq 'url') {
            $href = Prov-Esc ([string]$node.url)
            $name = Prov-Esc ([string]$node.name)
            [void]$script:ProvBmLines.Add("$pad<DT><A HREF=`"$href`">$name</A>")
        } elseif ($type -eq 'folder') {
            $name = Prov-Esc ([string]$node.name)
            # Eviter de re-wrapper "Barre de favoris" / "Bookmarks bar" en sous-dossier
            if ($name -match '(?i)^(barre\s+de\s+favoris|bookmarks?\s*bar|barra\s+de\s+favoritos)$') {
                foreach ($child in @($node.children)) {
                    & $writer $child $depth
                }
                return
            }
            [void]$script:ProvBmLines.Add("$pad<DT><H3>$name</H3>")
            [void]$script:ProvBmLines.Add("$pad<DL><p>")
            foreach ($child in @($node.children)) {
                & $writer $child ($depth + 1)
            }
            [void]$script:ProvBmLines.Add("$pad</DL><p>")
        }
    }

    # 1) Barre de favoris → attribut reconnu par Chrome a l'import
    try {
        $bar = $data.roots.bookmark_bar
        if ($bar) {
            [void]$lines.Add('    <DT><H3 PERSONAL_TOOLBAR_FOLDER="true">Bookmarks Bar</H3>')
            [void]$lines.Add('    <DL><p>')
            foreach ($child in @($bar.children)) {
                & $writer $child 2
            }
            [void]$lines.Add('    </DL><p>')
        }
    } catch {}

    # 2) Autres / sync → dossiers classiques (pas la barre)
    $otherMap = @{
        other  = 'Other Bookmarks'
        synced = 'Mobile Bookmarks'
    }
    foreach ($rootName in @('other', 'synced')) {
        try {
            $root = $data.roots.$rootName
            if (-not $root) { continue }
            $kids = @($root.children)
            if ($kids.Count -eq 0) { continue }
            $title = $otherMap[$rootName]
            [void]$lines.Add(("    <DT><H3>{0}</H3>" -f (Prov-Esc $title)))
            [void]$lines.Add('    <DL><p>')
            foreach ($child in $kids) {
                & $writer $child 2
            }
            [void]$lines.Add('    </DL><p>')
        } catch {}
    }

    [void]$lines.Add('</DL><p>')
    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($OutHtmlPath, ($lines -join "`r`n"), $utf8Bom)
    Remove-Variable -Name ProvBmLines -Scope Script -EA SilentlyContinue
    return (Test-Path -LiteralPath $OutHtmlPath)
}

function Export-ProvChromeBookmarks {
    <#
    Favoris Chrome — tous les profils Default / Profile N
      chrome-bookmarks.html (+ chrome-bookmarks-Profile_1.html …)
    #>
    param([string]$DestDir)

    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " Favoris Chrome " + ($UI.HLine * 36)) -ForegroundColor Magenta

    $profiles = @(Get-ProvChromeProfiles)
    if ($profiles.Count -eq 0) {
        Write-MigStatus -Name "Favoris" -Status SKIP -Detail "profil Chrome introuvable"
        Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
        return $false
    }

    Get-Process chrome -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
    Start-Sleep 1

    Write-Host ("  " + $UI.VLine + "  $($profiles.Count) profil(s) — export auto (pas d'UI Chrome)") -ForegroundColor Gray
    $ok = $false
    foreach ($pr in $profiles) {
        $bm = Join-Path $pr.Path "Bookmarks"
        if (-not (Test-Path -LiteralPath $bm)) { continue }
        $jsonOut = Join-Path $DestDir ("chrome-bookmarks{0}.json" -f $pr.FileSuffix)
        $htmlOut = Join-Path $DestDir ("chrome-bookmarks{0}.html" -f $pr.FileSuffix)
        if ($pr.Directory -eq 'Default') {
            $jsonOut = Join-Path $DestDir "chrome-bookmarks.json"
            $htmlOut = Join-Path $DestDir "chrome-bookmarks.html"
        }
        try {
            Copy-Item -LiteralPath $bm -Destination $jsonOut -Force
            $htmlOk = ConvertTo-ProvChromeBookmarksHtml -BookmarksJsonPath $bm -OutHtmlPath $htmlOut
            Write-MigStatus -Name ("Favoris ($($pr.DisplayName))") -Status OK -Detail $(if ($htmlOk) { (Split-Path $htmlOut -Leaf) } else { (Split-Path $jsonOut -Leaf) })
            $ok = $true
        } catch {
            Write-MigStatus -Name ("Favoris ($($pr.DisplayName))") -Status FAIL -Detail $_.Exception.Message
        }
    }

    if (-not $ok) {
        Write-MigStatus -Name "Favoris" -Status SKIP -Detail "aucun fichier Bookmarks"
    } else {
        Write-Host ("  " + $UI.VLine + "  Sur le neuf : chaque HTML s'importe (ou se fusionne) dans un profil") -ForegroundColor DarkGray
    }
    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
    return $ok
}

function Write-ProvMigrationManifest {
    param(
        [string]$DestDir,
        [object]$Identity,
        [bool]$HasPasswords,
        [bool]$HasBookmarks
    )
    $pwdFiles = @(Get-ChildItem -LiteralPath $DestDir -Filter "chrome-passwords*.csv" -File -EA SilentlyContinue | Sort-Object Name | ForEach-Object { $_.Name })
    $bmHtml   = @(Get-ChildItem -LiteralPath $DestDir -Filter "chrome-bookmarks*.html" -File -EA SilentlyContinue | Sort-Object Name | ForEach-Object { $_.Name })
    $bmJson   = @(Get-ChildItem -LiteralPath $DestDir -Filter "chrome-bookmarks*.json" -File -EA SilentlyContinue | Sort-Object Name | ForEach-Object { $_.Name })
    $manifest = [ordered]@{
        schema             = ("{0}-migration-v1" -f $ProvConfig.FilePrefix)
        folderName         = Split-Path $DestDir -Leaf
        prenom             = $Identity.Prenom
        nom                = $Identity.Nom
        email              = $Identity.Email
        sourcePc           = $env:COMPUTERNAME
        sourceUser         = $env:USERNAME
        exportedAt         = (Get-Date -Format "o")
        chromePasswords    = $(if ($pwdFiles.Count -gt 0) { $pwdFiles[0] } elseif ($HasPasswords) { "chrome-passwords.csv" } else { $null })
        chromePasswordsAll = $pwdFiles
        chromeBookmarks    = $(if ($bmHtml.Count -gt 0) { $bmHtml[0] } elseif ($bmJson.Count -gt 0) { $bmJson[0] } elseif ($HasBookmarks) { "chrome-bookmarks.html" } else { $null })
        chromeBookmarksAll = $bmHtml
        chromeBookmarksJson = $bmJson
        dirs               = @("Desktop", "Documents", "Downloads")
        note               = "Profile_N files map to Chrome folder 'Profile N' (underscore = space)"
    }
    $json = $manifest | ConvertTo-Json -Depth 5
    Set-Content -Path (Join-Path $DestDir ("{0}-migration.json" -f $ProvConfig.FilePrefix)) -Value $json -Encoding UTF8

    $readme = @"
$($ProvConfig.OrgName) migrationWin11 — $($Identity.Prenom) $($Identity.Nom)
Source : $env:COMPUTERNAME / $env:USERNAME
Date   : $(Get-Date -Format "yyyy-MM-dd HH:mm")

Contenu Chrome (multi-profil)
-----------------------------
- chrome-passwords.csv              → profil Default
- chrome-passwords-Profile_N.csv    → profil Chrome "Profile N"
- chrome-bookmarks.html / .json     → idem pour favoris
- chrome-bookmarks-Profile_N.html

Sur le NOUVEAU poste : deploy ouvre chaque profil Chrome correspondant.
CSV = CLAIR — supprimer apres import.
"@
    Set-Content (Join-Path $DestDir $ProvConfig.NoteFileName) -Value $readme -Encoding UTF8
}

# ============================================================================
#  MAIN
# ============================================================================

Set-ProvConsoleMaximized
Clear-Host
Write-Host ""
Write-Host ("  " + ("=" * 58)) -ForegroundColor DarkCyan
Write-Host ("   " + $ProvConfig.OrgName + "  ·  Migration ancien poste → cle USB") -ForegroundColor Cyan
Write-Host "   Selection interactive + migrationWin11_prenom_nom" -ForegroundColor DarkGray
Write-Host ("  " + ("=" * 58)) -ForegroundColor DarkCyan
Write-Host ""

if (Test-ProvIsAdmin) {
    Write-Host "  Session ADMIN — preferez le compte user sans elevation." -ForegroundColor Red
    $go = (Read-Host "     Continuer quand meme ? [o/N]").Trim().ToLowerInvariant()
    if ($go -notin @('o','oui','y','yes')) { Read-Host "  Entree"; exit 1 }
}

$identity = Read-ProvMigrationFolderName
$dest = Read-ProvUsbDestination -FolderName $identity.FolderName
if (-not $dest) { Read-Host "  Entree"; exit 1 }

$desktop = [Environment]::GetFolderPath("Desktop")
$docs    = [Environment]::GetFolderPath("MyDocuments")
$dl      = Join-Path $env:USERPROFILE "Downloads"

$deskItems = @(Get-ProvTopLevelItems -Path $desktop -SkipApps)
$docItems  = @(Get-ProvTopLevelItems -Path $docs)
$dlItems   = @(Get-ProvTopLevelItems -Path $dl)

Write-Host ""
Write-Host ("  Bureau detecte          : {0}  ($desktop)" -f $deskItems.Count) -ForegroundColor DarkGray
Write-Host ("  Documents detectes      : {0}  ($docs)" -f $docItems.Count) -ForegroundColor DarkGray
Write-Host ("  Telechargements detectes: {0}  ($dl)" -f $dlItems.Count) -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Selectionnez ce qu'il faut conserver (niveau 1 uniquement :" -ForegroundColor Gray
Write-Host "  un dossier = tout son contenu)." -ForegroundColor Gray
Start-Sleep -Seconds 1

$selDesk = @(Show-ProvCheckList -Title "BUREAU — cocher pour conserver" -Items $deskItems)
$selDocs = @(Show-ProvCheckList -Title "DOCUMENTS — cocher pour conserver" -Items $docItems)
$selDl   = @(Show-ProvCheckList -Title "TELECHARGEMENTS — cocher pour conserver" -Items $dlItems)

Write-Host ""
Write-Host "  Recap :" -ForegroundColor Cyan
Write-Host ("    Bureau          : {0} element(s)" -f $selDesk.Count)
Write-Host ("    Documents       : {0} element(s)" -f $selDocs.Count)
Write-Host ("    Telechargements : {0} element(s)" -f $selDl.Count)
Write-Host ("    Destination     : $dest") -ForegroundColor Yellow
Write-Host ""
$okGo = (Read-Host "     Lancer la copie ? [O/n]").Trim().ToLowerInvariant()
if ($okGo -in @('n','non','no')) {
    Write-MigStatus -Name "Migration" -Status SKIP -Detail "annule"
    Read-Host "  Entree"
    exit 0
}

Write-Host ""
Write-Host "  --- Copie ---" -ForegroundColor Cyan
[void](Copy-ProvSelectedItems -Label "Bureau" -DestSubDir (Join-Path $dest "Desktop") -Items $selDesk)
[void](Copy-ProvSelectedItems -Label "Documents" -DestSubDir (Join-Path $dest "Documents") -Items $selDocs)
[void](Copy-ProvSelectedItems -Label "Telechargements" -DestSubDir (Join-Path $dest "Downloads") -Items $selDl)

Write-Host ""
Write-Host "  --- Mots de passe Chrome ---" -ForegroundColor Cyan
$hasPwd = Export-ProvChromePasswords -DestDir $dest

Write-Host ""
Write-Host "  --- Favoris Chrome ---" -ForegroundColor Cyan
$hasBm = Export-ProvChromeBookmarks -DestDir $dest

Write-ProvMigrationManifest -DestDir $dest -Identity $identity -HasPasswords $hasPwd -HasBookmarks $hasBm

try { Start-Process explorer.exe -ArgumentList $dest } catch {}

Write-Host ""
Write-Host ("  " + ("=" * 58)) -ForegroundColor DarkCyan
Write-Host "   Migration prete : $dest" -ForegroundColor Green
Write-Host "   Sur le neuf : deploy proposera l'import au demarrage." -ForegroundColor Gray
Write-Host ("  " + ("=" * 58)) -ForegroundColor DarkCyan
Write-Host ""
Read-Host "  Entree pour fermer"
