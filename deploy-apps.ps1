# ============================================================================
#  Provisioning - Deploiement applications (Win11 Pro / Intune / M365 Business Premium)
#  Fichier : deploy-apps.ps1
#  Usage   : start.bat  |  -ModulePicker  |  -Modules startup,chrome-default  |  -ListModules
#
#  IMPORTANT : ce fichier DOIT rester en UTF-8 avec BOM (Windows PowerShell 5.1).
#  Valeurs propres a l'organisation -> config.local.psd1 (modele : config.example.psd1).
#
#  Hors-ligne / contournements :
#    Le script cherche d'ABORD sur la cle USB (E:\scripts\offline, usb_pro\downloads…),
#    puis a cote du .ps1 (Bureau\scripts\offline, etc.).
#    Assets : AnyDesk.exe, ODT, apps metier declarees dans la config…
# ============================================================================

#Requires -RunAsAdministrator
param(
    # Ex. -Modules startup,chrome-default  |  -Modules perso  |  -ModulePicker
    [string[]]$Modules = @(),
    [switch]$ModulePicker,
    [switch]$ListModules
)

$ErrorActionPreference = "Continue"
$ProgressPreference    = "SilentlyContinue"

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
    chcp 65001 | Out-Null
} catch {}

# --- Glyphes UI (codes Unicode) : parse-safe meme si le BOM est perdu ---
$UI = @{
    Check   = [string][char]0x2714   # heavy check
    Dot     = [string][char]0x25CF   # black circle
    Cross   = [string][char]0x2716   # heavy multiplication x
    Play    = [string][char]0x25B8   # black right-pointing small triangle
    Info    = [string][char]0x2139   # information source
    DotMid  = [string][char]0x00B7   # middle dot
    Ellipsis= [string][char]0x2026   # horizontal ellipsis
    EmDash  = [string][char]0x2014   # em dash
    Arrow   = [string][char]0x2192   # rightwards arrow
    HLine   = [string][char]0x2500   # box drawings light horizontal
    TL      = [string][char]0x250C   # top-left
    BL      = [string][char]0x2514   # bottom-left
    VLine   = [string][char]0x2502   # vertical
    DLine   = [string][char]0x2550   # double horizontal
    BoxTL   = [string][char]0x2554
    BoxTR   = [string][char]0x2557
    BoxBL   = [string][char]0x255A
    BoxBR   = [string][char]0x255D
    BoxV    = [string][char]0x2551
    BoxML   = [string][char]0x2560
    BoxMR   = [string][char]0x2563
}
$UI.Spinner = @(
    [string][char]0x280B, [string][char]0x2819, [string][char]0x2839, [string][char]0x2838,
    [string][char]0x283C, [string][char]0x2834, [string][char]0x2826, [string][char]0x2827,
    [string][char]0x2807, [string][char]0x280F
)
$UI.Rule = ($UI.DLine * 58)
$UI.RuleSm = ($UI.HLine * 54)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
$LogDate = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$script:Stats = @{ OK = 0; Skip = 0; Fail = 0 }
$script:LogEnabled = $true

# Codes winget : deja installe / pas de MAJ
$WINGET_ALREADY = @(-1978335189)  # UPDATE_NOT_APPLICABLE seulement — PAS -1978335212 (package introuvable)


# ============================================================================
#  CONFIG — valeurs propres a l'organisation : config.local.psd1 (non versionne)
#  Modele commente : config.example.psd1
# ============================================================================
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

# Apps winget (hors Chrome / AnyDesk) — aussi utilisees en mode -Modules apps
$script:AppsWinget = @($ProvConfig.WingetApps)

# Identite collectee pendant le deploiement (remplie par les prompts)
$script:ProvPrenom = ""
$script:ProvNom = ""
$script:ProvEmail = ""
$script:ProvComputerName = $env:COMPUTERNAME
$script:ReportPaths = @()
$script:AnyDeskId = ""
$script:ProvMigrationRoot = $null


function Test-ProvPathWritable {
    <# Teste si on peut ecrire dans ce dossier. Ne cree le dossier que si -Create. #>
    param(
        [string]$Dir,
        [switch]$Create
    )
    if ([string]::IsNullOrWhiteSpace($Dir)) { return $false }
    try {
        if (-not (Test-Path -LiteralPath $Dir)) {
            if (-not $Create) { return $false }
            New-Item -ItemType Directory -Path $Dir -Force -ErrorAction Stop | Out-Null
        }
        $probe = Join-Path $Dir ("{0}-write-test-{1}.tmp" -f $ProvConfig.FilePrefix, [guid]::NewGuid().ToString("N"))
        [System.IO.File]::WriteAllText($probe, "ok")
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return $true
    } catch {
        return $false
    }
}

function Get-ProvReportDir {
    <#
    Documents\<ReportFolderName> — logs + inventaire (copie locale poste).
    Miroir USB : Sync-ProvReportsToUsb → scripts\<UsbReportsFolder>\<PC>-<stamp>\
    #>
    $docs = [Environment]::GetFolderPath("MyDocuments")
    $candidates = @(
        (Join-Path $docs $ProvConfig.ReportFolderName),
        (Join-Path $env:LOCALAPPDATA $ProvConfig.ReportFolderName),
        (Join-Path $env:TEMP $ProvConfig.ReportFolderName)
    )
    foreach ($d in $candidates) {
        if (Test-ProvPathWritable -Dir $d -Create) { return $d }
    }
    return $null
}

function Resolve-ProvLogFile {
    <# Journal principal dans Documents (miroir USB en fin de deploy). #>
    param([string]$PreferredDir, [string]$Stamp)
    $name = "install-log-$Stamp.txt"
    $candidates = @(
        $PreferredDir,
        (Join-Path $env:TEMP $ProvConfig.WorkFolderName),
        (Join-Path $env:LOCALAPPDATA $ProvConfig.WorkFolderName),
        $env:TEMP
    ) | Where-Object { $_ } | Select-Object -Unique

    foreach ($dir in $candidates) {
        if (Test-ProvPathWritable -Dir $dir -Create) {
            return [pscustomobject]@{
                Path     = (Join-Path $dir $name)
                Dir      = $dir
                Fallback = ($PreferredDir -and ($dir -ne $PreferredDir))
            }
        }
    }
    return $null
}

$script:ProvReportDir = Get-ProvReportDir
$script:ProvUsbReportDir = $null
$script:ProvStatePath = $null
$logInfo = Resolve-ProvLogFile -PreferredDir $script:ProvReportDir -Stamp $LogDate
if ($logInfo) {
    $LogFile = $logInfo.Path
    $script:LogDir = $logInfo.Dir
    $script:LogFallback = [bool]$logInfo.Fallback
    if ($script:ProvReportDir -and ($script:LogDir -eq $script:ProvReportDir)) {
        $script:LogFallback = $false
    }
} else {
    $LogFile = $null
    $script:LogEnabled = $false
    $script:LogFallback = $true
}

function Get-ProvDeployStatePath {
    <# JSON local poste — pas sur la cle (infos identite entre deux runs). #>
    if ($script:ProvStatePath) { return $script:ProvStatePath }
    $dirs = @(
        $script:ProvReportDir,
        (Join-Path $env:LOCALAPPDATA $ProvConfig.WorkFolderName),
        (Join-Path $env:TEMP $ProvConfig.WorkFolderName)
    ) | Where-Object { $_ }
    foreach ($d in $dirs) {
        if (Test-ProvPathWritable -Dir $d -Create) {
            $script:ProvStatePath = Join-Path $d ("{0}-deploy-state.json" -f $ProvConfig.FilePrefix)
            return $script:ProvStatePath
        }
    }
    return $null
}

function Read-ProvDeployState {
    $path = Get-ProvDeployStatePath
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        if (-not $obj) { return $null }
        # Ignore si plus vieux que 14 jours
        if ($obj.UpdatedAt) {
            try {
                $when = [datetime]::Parse($obj.UpdatedAt)
                if ((Get-Date) - $when -gt [timespan]::FromDays(14)) { return $null }
            } catch {}
        }
        # Ignore si un deploy precedent a marque Completed
        if ($obj.Completed -eq $true) { return $null }
        return $obj
    } catch {
        return $null
    }
}

function Save-ProvDeployState {
    <# Persiste identite de base (PAS de mots de passe). #>
    param([hashtable]$Patch)
    $path = Get-ProvDeployStatePath
    if (-not $path) { return }

    $cur = $null
    try {
        if (Test-Path -LiteralPath $path) {
            $cur = (Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json)
        }
    } catch { $cur = $null }

    $h = @{
        Version      = 1
        Completed    = $false
        UpdatedAt    = (Get-Date -Format "o")
        HostAtSave   = $env:COMPUTERNAME
        ComputerName = $(if ($cur -and $cur.ComputerName) { [string]$cur.ComputerName } else { "" })
        Prenom       = $(if ($cur -and $cur.Prenom) { [string]$cur.Prenom } else { "" })
        Nom          = $(if ($cur -and $cur.Nom) { [string]$cur.Nom } else { "" })
        Email        = $(if ($cur -and $cur.Email) { [string]$cur.Email } else { "" })
    }
    if ($Patch) {
        foreach ($k in $Patch.Keys) {
            $h[$k] = $Patch[$k]
        }
    }
    # Toujours refleter les vars script courantes si non vides
    if ($script:ProvComputerName) { $h.ComputerName = [string]$script:ProvComputerName }
    if ($script:ProvPrenom) { $h.Prenom = [string]$script:ProvPrenom }
    if ($script:ProvNom) { $h.Nom = [string]$script:ProvNom }
    if ($script:ProvEmail) { $h.Email = [string]$script:ProvEmail }
    $h.UpdatedAt = (Get-Date -Format "o")
    $h.HostAtSave = $env:COMPUTERNAME
    $h.Completed = $false

    try {
        ($h | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $path -Encoding UTF8 -Force
        Write-Log "State sauve: $path" -Level INFO
    } catch {
        Write-Log "State save KO: $($_.Exception.Message)" -Level WARN
    }
}

function Clear-ProvDeployState {
    param([switch]$MarkCompletedOnly)
    $path = Get-ProvDeployStatePath
    if (-not $path) { return }
    try {
        if ($MarkCompletedOnly -and (Test-Path -LiteralPath $path)) {
            $obj = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
            $obj | Add-Member -NotePropertyName Completed -NotePropertyValue $true -Force
            $obj | Add-Member -NotePropertyName CompletedAt -NotePropertyValue (Get-Date -Format "o") -Force
            ($obj | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $path -Encoding UTF8 -Force
            return
        }
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force -ErrorAction Stop
            Write-Log "State efface: $path" -Level INFO
        }
    } catch {
        Write-Log "State clear KO: $($_.Exception.Message)" -Level WARN
    }
}

# ============================================================================
#  UI — terminal "ops" (style hacker, glyphes via [char] uniquement)
# ============================================================================

# ============================================================================
#  UI
# ============================================================================

function Write-Log {
    param([string]$Message, [ValidateSet("INFO","OK","WARN","ERR","STEP")][string]$Level = "INFO")
    if (-not $script:LogEnabled -or -not $LogFile) { return }
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try {
        Add-Content -Path $LogFile -Value "[$stamp] [$Level] $Message" -Encoding UTF8 -ErrorAction Stop
    } catch {
        $script:LogEnabled = $false
    }
}

function Write-Banner {
    Clear-Host
    Write-Host ""
    Write-Host ("  " + $UI.Rule) -ForegroundColor DarkCyan
    Write-Host ("   " + $ProvConfig.OrgName + "  " + $UI.DotMid + "  Deploiement poste de travail") -ForegroundColor Cyan
    Write-Host ("   Microsoft 365 Business Premium  " + $UI.DotMid + "  Windows 11") -ForegroundColor DarkGray
    Write-Host ("  " + $UI.Rule) -ForegroundColor DarkCyan
    Write-Host ""
}

function Set-ProvConsoleMaximized {
    <# Plein ecran (maximise) la console admin — une fois les droits OK. #>
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
        if ($hwnd -ne [IntPtr]::Zero) {
            [void][ProvConsoleWin]::ShowWindow($hwnd, [ProvConsoleWin]::SW_MAXIMIZE)
        }
    } catch {
        try {
            # Fallback : buffer + fenetre console au max
            $ui = $Host.UI.RawUI
            $ui.WindowSize = $ui.MaxWindowSize
        } catch {}
    }
}

function Read-ProvYesNo {
    <#
    Prompt bloquant O/N. Ne retourne que quand l'operateur repond.
    Default Yes => Entree = Oui (affiche O/n).
    #>
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [ValidateSet('Yes','No')][string]$Default = 'Yes'
    )
    $hint = if ($Default -eq 'Yes') { 'O/n' } else { 'o/N' }
    while ($true) {
        $ans = (Read-Host "     $Prompt [$hint]").Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($ans)) {
            return ($Default -eq 'Yes')
        }
        if ($ans -in @('o', 'oui', 'y', 'yes', 'ok')) { return $true }
        if ($ans -in @('n', 'non', 'no')) { return $false }
        Write-Host "     Repondre O (oui / OK) ou N (non) — le script attend." -ForegroundColor Yellow
    }
}

function Write-Section {
    param([string]$Title, [string]$Step)
    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " [$Step] " + ($UI.HLine * 45)) -ForegroundColor DarkCyan
    Write-Host ("  " + $UI.VLine + "  $Title") -ForegroundColor White
    Write-Host ("  " + $UI.BL + ($UI.HLine * 54)) -ForegroundColor DarkCyan
    Write-Host ""
    Write-Log $Title -Level STEP
}

function Write-Status {
    param(
        [string]$Name,
        [ValidateSet("OK","SKIP","FAIL","RUN","INFO")][string]$Status,
        [string]$Detail = "",
        [switch]$NoCount
    )
    if (-not $NoCount) {
        switch ($Status) {
            "OK"   { $script:Stats.OK++ }
            "SKIP" { $script:Stats.Skip++ }
            "FAIL" { $script:Stats.Fail++ }
        }
    }
    $icon = switch ($Status) {
        "OK"   { $UI.Check }
        "SKIP" { $UI.Dot }
        "FAIL" { $UI.Cross }
        "RUN"  { $UI.Play }
        "INFO" { $UI.Info }
    }
    $color = switch ($Status) {
        "OK"   { "Green" }
        "SKIP" { "DarkYellow" }
        "FAIL" { "Red" }
        "RUN"  { "Cyan" }
        "INFO" { "Gray" }
    }
    $tag = switch ($Status) {
        "OK"   { "  OK      " }
        "SKIP" { "  DEJA    " }
        "FAIL" { "  ECHEC   " }
        "RUN"  { "  " + $UI.Ellipsis + "       " }
        "INFO" { "  " + $UI.DotMid + "       " }
    }
    Write-Host "    $icon " -ForegroundColor $color -NoNewline
    Write-Host ($Name.PadRight(32)) -NoNewline
    Write-Host $tag -ForegroundColor $color -NoNewline
    if ($Detail) { Write-Host $Detail -ForegroundColor DarkGray } else { Write-Host "" }
    $lvl = switch ($Status) { "OK"{"OK"} "SKIP"{"WARN"} "FAIL"{"ERR"} default {"INFO"} }
    Write-Log ("$Name $($UI.EmDash) $Status $(if($Detail){"($Detail)"})") -Level $lvl
}

function Write-Summary {
    Write-Host ""
    Write-Host ("  " + $UI.BoxTL + ($UI.DLine * 52) + $UI.BoxTR) -ForegroundColor DarkCyan
    Write-Host ("  " + $UI.BoxV + "              RESUME DU DEPLOIEMENT                 " + $UI.BoxV) -ForegroundColor Cyan
    Write-Host ("  " + $UI.BoxML + ($UI.DLine * 52) + $UI.BoxMR) -ForegroundColor DarkCyan
    Write-Host ("  " + $UI.BoxV + ("  $($UI.Check)  Reussis   : {0,-5}                            " -f $script:Stats.OK) + $UI.BoxV) -ForegroundColor Green
    Write-Host ("  " + $UI.BoxV + ("  $($UI.Dot)  Deja OK   : {0,-5}                            " -f $script:Stats.Skip) + $UI.BoxV) -ForegroundColor DarkYellow
    $fc = if ($script:Stats.Fail -gt 0) { "Red" } else { "Green" }
    Write-Host ("  " + $UI.BoxV + ("  $($UI.Cross)  Echecs    : {0,-5}                            " -f $script:Stats.Fail) + $UI.BoxV) -ForegroundColor $fc
    Write-Host ("  " + $UI.BoxML + ($UI.DLine * 52) + $UI.BoxMR) -ForegroundColor DarkCyan
    if ($script:LogEnabled -and $LogFile) {
        Write-Host ("  " + $UI.BoxV + "  Journal : $LogFile") -ForegroundColor Gray
    } else {
        Write-Host ("  " + $UI.BoxV + "  Journal : (desactive)") -ForegroundColor DarkGray
    }
    Write-Host ("  " + $UI.BoxBL + ($UI.DLine * 52) + $UI.BoxBR) -ForegroundColor DarkCyan
    Write-Host ""
}

function Wait-ProcessWithSpinner {
    param([System.Diagnostics.Process]$Process, [string]$Label = "Traitement")
    $i = 0
    while (-not $Process.HasExited) {
        $frame = $UI.Spinner[$i % $UI.Spinner.Count]
        Write-Host ("`r    {0}  {1}{2}" -f $frame, $Label, $UI.Ellipsis) -ForegroundColor Cyan -NoNewline
        Start-Sleep -Milliseconds 100
        $i++
    }
    Write-Host "`r$(' ' * 72)`r" -NoNewline
}

# ============================================================================
#  WINGET
# ============================================================================

function Test-WingetAvailable {
    return $null -ne (Get-Command winget -ErrorAction SilentlyContinue)
}

function Invoke-Winget {
    param([string]$Arguments)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = "winget"
    $psi.Arguments              = $Arguments
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    $out = $p.StandardOutput.ReadToEnd()
    $err = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
    return [pscustomobject]@{ ExitCode = $p.ExitCode; StdOut = $out; StdErr = $err }
}

function Get-WingetExitMeaning {
    param([int]$Code)
    switch ($Code) {
        0            { return "OK" }
        -1978335231  { return "erreur interne winget" }
        -1978335230  { return "arguments winget invalides" }
        -1978335229  { return "commande winget echouee" }
        -1978335222  { return "index winget corrompu" }
        -1978335217  { return "donnees source manquantes" }
        -1978335216  { return "aucun installateur applicable" }
        -1978335215  { return "hash mismatch (manifest winget obsolete)" }  # 0x8A150011
        -1978335212  { return "package introuvable (winget)" }
        -1978335189  { return "deja installe / pas de MAJ" }                 # INSTALL_ALREADY_INSTALLED
        -1978335187  { return "controle securite / hash bloque (admin)" }    # INSTALLER_SECURITY_CHECK_FAILED
        -1978335174  { return "bloque par strategie (GPO/Intune)" }
        -1978335167  { return "erreur reseau" }
        -1978335138  { return "certificat TLS pin mismatch (proxy/CDN/cache)" } # PINNED_CERTIFICATE_MISMATCH 0x8A15005E
        -1978335135  { return "package deja installe (winget)" }             # PACKAGE_ALREADY_INSTALLED 0x8A150061
        default      { return "code $Code" }
    }
}

function Invoke-WingetInstall {
    param(
        [string]$Id,
        [string]$Label,
        [string]$ExtraArgs = ""
    )
    # --source winget : evite msstore (pin TLS frequent avec proxy/SSL inspection)
    $arg = "install --id `"$Id`" --exact --source winget --silent --accept-package-agreements --accept-source-agreements --disable-interactivity --force"
    if ($ExtraArgs) { $arg = "$arg $ExtraArgs" }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = "winget"
    $psi.Arguments              = $arg
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    $outSb = New-Object System.Text.StringBuilder
    $errSb = New-Object System.Text.StringBuilder
    $outEv = Register-ObjectEvent -InputObject $p -EventName OutputDataReceived -MessageData $outSb -Action {
        if ($EventArgs.Data) { [void]$Event.MessageData.AppendLine($EventArgs.Data) }
    }
    $errEv = Register-ObjectEvent -InputObject $p -EventName ErrorDataReceived -MessageData $errSb -Action {
        if ($EventArgs.Data) { [void]$Event.MessageData.AppendLine($EventArgs.Data) }
    }
    $p.BeginOutputReadLine()
    $p.BeginErrorReadLine()
    Wait-ProcessWithSpinner -Process $p -Label $Label
    $p.WaitForExit()
    Unregister-Event -SourceIdentifier $outEv.Name -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier $errEv.Name -ErrorAction SilentlyContinue
    Get-Job | Where-Object { $_.Id -in @($outEv.Id, $errEv.Id) } | Remove-Job -Force -ErrorAction SilentlyContinue
    return [pscustomobject]@{
        ExitCode = $p.ExitCode
        StdOut   = $outSb.ToString()
        StdErr   = $errSb.ToString()
    }
}

function Initialize-ProvWingetNetwork {
    <#
      Contourne pin TLS winget (proxy / SSL inspection / antivirus MITM).
      Admin settings + policy HKLM + preferer source community.
    #>
    Write-Status -Name "winget reseau" -Status RUN -Detail "bypass pin TLS + hash override" -NoCount
    try { & winget settings --enable BypassCertificatePinningForMicrosoftStore 2>$null | Out-Null } catch {}
    try { & winget settings --enable InstallerHashOverride 2>$null | Out-Null } catch {}
    try {
        $key = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppInstaller"
        if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
        New-ItemProperty -Path $key -Name "EnableBypassCertificatePinningForMicrosoftStore" -PropertyType DWord -Value 1 -Force -ErrorAction SilentlyContinue | Out-Null
        New-ItemProperty -Path $key -Name "EnableInstallerHashOverride" -PropertyType DWord -Value 1 -Force -ErrorAction SilentlyContinue | Out-Null
    } catch {
        Write-Log "AppInstaller policy : $($_.Exception.Message)" -Level WARN
    }
    try { & winget source update --disable-interactivity 2>$null | Out-Null } catch {}
    Write-Status -Name "winget reseau" -Status OK -Detail "pin bypass + source winget" -NoCount
}

function Test-ProvAppOnDisk {
    param([string]$Id)
    $paths = switch -Regex ($Id) {
        '^Mozilla\.Firefox' {
            @(
                "${env:ProgramFiles}\Mozilla Firefox\firefox.exe",
                "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"
            )
        }
        '^Discord\.' {
            @(
                "$env:LOCALAPPDATA\Discord\Update.exe",
                "${env:LOCALAPPDATA}\Discord\app-*\Discord.exe"
            )
        }
        '^7zip\.' {
            @(
                "${env:ProgramFiles}\7-Zip\7z.exe",
                "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
            )
        }
        '^Bitwarden\.' {
            @(
                "${env:ProgramFiles}\Bitwarden\Bitwarden.exe",
                "$env:LOCALAPPDATA\Programs\Bitwarden\Bitwarden.exe"
            )
        }
        'Notepad\+\+' {
            @(
                "${env:ProgramFiles}\Notepad++\notepad++.exe",
                "${env:ProgramFiles(x86)}\Notepad++\notepad++.exe"
            )
        }
        'Temurin\.17\.JRE|JavaRuntimeEnvironment' {
            @(
                "${env:ProgramFiles}\Java\*\bin\java.exe",
                "${env:ProgramFiles}\Eclipse Adoptium\*\bin\java.exe",
                "${env:ProgramFiles}\Microsoft\jdk-*\bin\java.exe"
            )
        }
        default { @() }
    }
    foreach ($pat in $paths) {
        if ($pat -match '[\*\?]') {
            $hit = Get-Item -Path $pat -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($hit) { return $true }
        } elseif (Test-Path -LiteralPath $pat) {
            return $true
        }
    }
    return $false
}

function Get-ProvGithubLatestAssetUrl {
    param(
        [Parameter(Mandatory)][string]$Repo,
        [Parameter(Mandatory)][string]$NameMatch
    )
    try {
        $api = "https://api.github.com/repos/$Repo/releases/latest"
        $headers = @{ "User-Agent" = [string]$ProvConfig.UserAgent; "Accept" = "application/vnd.github+json" }
        $rel = Invoke-RestMethod -Uri $api -Headers $headers -TimeoutSec 45 -ErrorAction Stop
        $asset = @($rel.assets) | Where-Object { $_.name -match $NameMatch } | Select-Object -First 1
        if ($asset -and $asset.browser_download_url) { return [string]$asset.browser_download_url }
    } catch {
        Write-Log "GitHub latest $Repo : $($_.Exception.Message)" -Level WARN
    }
    return $null
}

function Get-ProvDirectInstallerSpec {
    param([string]$Id)
    switch ($Id) {
        'Mozilla.Firefox' {
            return @{
                Offline = @('FirefoxSetup.exe', 'Firefox Setup*.exe', 'Firefox*.exe')
                Urls    = @(
                    'https://download.mozilla.org/?product=firefox-latest-ssl&os=win64&lang=fr',
                    'https://download.mozilla.org/?product=firefox-latest&os=win64&lang=fr'
                )
                File = 'FirefoxSetup.exe'; Args = '/S'
            }
        }
        'Discord.Discord' {
            return @{
                Offline = @('DiscordSetup.exe', 'Discord*.exe')
                Urls    = @(
                    'https://discord.com/api/downloads/distributions/app/installers/latest?channel=stable&platform=win&arch=x64'
                )
                File = 'DiscordSetup.exe'; Args = '-s'
            }
        }
        '7zip.7zip' {
            $gh = Get-ProvGithubLatestAssetUrl -Repo 'ip7z/7zip' -NameMatch '7z\d+-x64\.exe$'
            $urls = @()
            if ($gh) { $urls += $gh }
            $urls += @(
                'https://www.7-zip.org/a/7z2501-x64.exe',
                'https://www.7-zip.org/a/7z2409-x64.exe',
                'https://www.7-zip.org/a/7z2408-x64.exe'
            )
            return @{ Offline = @('7z*-x64.exe', '7z*.exe'); Urls = $urls; File = '7z-x64.exe'; Args = '/S' }
        }
        'Bitwarden.Bitwarden' {
            return @{
                Offline = @('Bitwarden-Installer.exe', 'Bitwarden*.exe')
                Urls    = @(
                    'https://vault.bitwarden.com/download/?app=desktop&platform=windows&variant=exe',
                    'https://bitwarden.com/download/?app=desktop&platform=windows'
                )
                File = 'Bitwarden-Installer.exe'; Args = '/S'
            }
        }
        'Notepad++.Notepad++' {
            $gh = Get-ProvGithubLatestAssetUrl -Repo 'notepad-plus-plus/notepad-plus-plus' -NameMatch 'Installer\.x64\.exe$'
            $urls = @(); if ($gh) { $urls += $gh }
            return @{
                Offline = @('npp*.Installer.x64.exe', 'npp*.exe')
                Urls = $urls; File = 'npp-Installer.x64.exe'; Args = '/S'
            }
        }
        'EclipseAdoptium.Temurin.17.JRE' {
            $gh = Get-ProvGithubLatestAssetUrl -Repo 'adoptium/temurin17-binaries' -NameMatch 'OpenJDK17U-jre_x64_windows_hotspot_.*\.msi$'
            $urls = @(); if ($gh) { $urls += $gh }
            return @{
                Offline = @('OpenJDK17*jre*.msi', 'temurin*17*jre*.msi')
                Urls = $urls; File = 'temurin17-jre.msi'; Args = 'MSI'
            }
        }
        default { return $null }
    }
}

function Find-ProvOfflineGlob {
    param([string[]]$Patterns)
    foreach ($root in @(Get-ProvOfflineDirsOnly)) {
        foreach ($pat in $Patterns) {
            if ([string]::IsNullOrWhiteSpace($pat)) { continue }
            if ($pat -match '[\*\?]') {
                $hit = Get-ChildItem -LiteralPath $root -Filter $pat -File -ErrorAction SilentlyContinue |
                    Select-Object -First 1
                if ($hit) { return $hit.FullName }
            } else {
                $p = Join-Path $root $pat
                if (Test-Path -LiteralPath $p -PathType Leaf) { return $p }
            }
        }
    }
    $exact = @($Patterns | Where-Object { $_ -and ($_ -notmatch '[\*\?]') })
    if ($exact.Count -gt 0) {
        $f = Find-ProvOfflineFile -Names $exact
        if ($f) { return $f }
    }
    return $null
}

function Install-ProvDirectVendor {
    <# DL editeur via curl (store CA Windows, sans pin winget) + install silencieux. #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Id
    )
    if (Test-ProvAppOnDisk -Id $Id) {
        Write-Status -Name $Name -Status SKIP -Detail "deja present (disque)"
        return "SKIP"
    }
    $spec = Get-ProvDirectInstallerSpec -Id $Id
    if (-not $spec) {
        return "FAIL"
    }
    Write-Status -Name $Name -Status RUN -Detail "fallback direct (hors winget)" -NoCount

    $installer = $null
    try { $installer = Find-ProvOfflineGlob -Patterns $spec.Offline } catch {}
    if ($installer) {
        Write-Status -Name $Name -Status INFO -Detail "offline: $installer" -NoCount
    } else {
        $tmpDir = Join-Path $env:TEMP ("{0}-direct" -f $ProvConfig.WorkFolderName)
        New-Item -ItemType Directory -Path $tmpDir -Force -EA SilentlyContinue | Out-Null
        $out = Join-Path $tmpDir $spec.File
        $okDl = $false
        foreach ($u in @($spec.Urls)) {
            if (-not $u) { continue }
            Write-Status -Name $Name -Status RUN -Detail "telechargement editeur..." -NoCount
            if (Save-ProvDownload -Url $u -OutFile $out) {
                if ((Test-Path $out) -and ((Get-Item $out).Length -gt 500KB)) {
                    $installer = $out; $okDl = $true; break
                }
            }
        }
        if (-not $okDl) {
            Write-Status -Name $Name -Status FAIL -Detail "telechargement editeur impossible"
            return "FAIL"
        }
    }

    try {
        $ext = [IO.Path]::GetExtension($installer).ToLowerInvariant()
        if ($ext -eq '.msi' -or $spec.Args -eq 'MSI') {
            $p = Start-Process -FilePath "msiexec.exe" -ArgumentList @(
                "/i", "`"$installer`"", "/qn", "/norestart", "ALLUSERS=1"
            ) -Wait -PassThru -WindowStyle Hidden
        } else {
            $p = Start-Process -FilePath $installer -ArgumentList $spec.Args -Wait -PassThru -WindowStyle Hidden
        }
        Write-Log "Direct install $Id exit=$($p.ExitCode) file=$installer" -Level INFO
        Start-Sleep -Seconds 2
        if (Test-ProvAppOnDisk -Id $Id) {
            Write-Status -Name $Name -Status OK -Detail "installe (direct editeur)"
            return "OK"
        }
        if ($p.ExitCode -in @(0, 3010)) {
            Start-Sleep -Seconds 4
            if (Test-ProvAppOnDisk -Id $Id) {
                Write-Status -Name $Name -Status OK -Detail "installe (direct, delai)"
                return "OK"
            }
            if ($Id -match 'Discord' -and $p.ExitCode -eq 0) {
                Write-Status -Name $Name -Status OK -Detail "installe (Discord silent)"
                return "OK"
            }
        }
        Write-Status -Name $Name -Status FAIL -Detail ("direct exit $($p.ExitCode)")
        return "FAIL"
    } catch {
        Write-Status -Name $Name -Status FAIL -Detail $_.Exception.Message
        return "FAIL"
    }
}

function Test-WingetPackageInstalled {
    param([string]$Id)
    if (Test-ProvAppOnDisk -Id $Id) { return $true }
    $r = Invoke-Winget -Arguments "list --id `"$Id`" --exact --accept-source-agreements --disable-interactivity"
    if ($r.ExitCode -ne 0) { return $false }
    if ($r.StdOut -match "No installed package|No package found|Aucun package") { return $false }
    $lines = $r.StdOut -split "`r?`n" | Where-Object { $_ -match [regex]::Escape($Id) -and $_ -notmatch '^\s*Name\s+Id' }
    return [bool]$lines
}

function Install-WingetApp {
    param([string]$Name, [string]$Id, [int]$Retries = 1, [string]$ExtraArgs = "")

    Write-Status -Name $Name -Status RUN -Detail $Id -NoCount

    if (Test-WingetPackageInstalled -Id $Id) {
        Write-Status -Name $Name -Status SKIP -Detail "deja present"
        return "SKIP"
    }

    $lastCode = -1
    $lastMeaning = ""
    for ($attempt = 1; $attempt -le ($Retries + 1); $attempt++) {
        $label = if ($attempt -eq 1) { "Installation $Name" } else { "Nouvelle tentative ($attempt) $($UI.EmDash) $Name" }
        $result = Invoke-WingetInstall -Id $Id -Label $label -ExtraArgs $ExtraArgs
        $lastCode = $result.ExitCode
        $lastMeaning = Get-WingetExitMeaning -Code $lastCode
        Write-Log "winget $Id attempt=$attempt exit=$lastCode ($lastMeaning)" -Level INFO
        if ($result.StdOut) { Write-Log ($result.StdOut.Substring(0, [Math]::Min(800, $result.StdOut.Length))) -Level INFO }

        if ($lastCode -eq 0) {
            Write-Status -Name $Name -Status OK
            return "OK"
        }
        if (($WINGET_ALREADY -contains $lastCode) -or ($result.StdOut -match "already installed|deja install")) {
            Write-Status -Name $Name -Status SKIP -Detail "deja present"
            return "SKIP"
        }
        if (Test-ProvAppOnDisk -Id $Id) {
            Write-Status -Name $Name -Status OK -Detail "present sur disque"
            return "OK"
        }
        if (($lastCode -eq -1978335215) -and ($ExtraArgs -notmatch "ignore-security-hash")) {
            Write-Status -Name $Name -Status INFO -Detail "hash mismatch -> --ignore-security-hash" -NoCount
            $result2 = Invoke-WingetInstall -Id $Id -Label "$Name (ignore hash)" -ExtraArgs "$ExtraArgs --ignore-security-hash"
            Write-Log "winget $Id ignore-hash exit=$($result2.ExitCode)" -Level INFO
            if ($result2.ExitCode -eq 0 -or (Test-ProvAppOnDisk -Id $Id)) {
                Write-Status -Name $Name -Status OK -Detail "installe (ignore-security-hash)"
                return "OK"
            }
            $lastCode = $result2.ExitCode
            $lastMeaning = Get-WingetExitMeaning -Code $lastCode
        }
        if (($lastCode -eq -1978335138) -and ($attempt -eq 1)) {
            Write-Status -Name $Name -Status INFO -Detail "cert pin -> re-init winget + retry" -NoCount
            Initialize-ProvWingetNetwork
        }
        if ($attempt -le $Retries) { Start-Sleep -Seconds 2 }
    }

    # Fallback : DL editeur (curl = CA Windows, sans pin winget)
    if (Get-ProvDirectInstallerSpec -Id $Id) {
        Write-Status -Name $Name -Status INFO -Detail ("winget KO ($lastMeaning) " + $UI.Arrow + " direct editeur") -NoCount
        $st = Install-ProvDirectVendor -Name $Name -Id $Id
        if ($st -eq "OK" -or $st -eq "SKIP") { return $st }
    }

    Write-Status -Name $Name -Status FAIL -Detail $lastMeaning
    return "FAIL"
}

function Install-JavaRuntime {
    # Requis pour le logiciel metier (config LobApp)
    Write-Status -Name "Java Runtime" -Status INFO -Detail "requis pour le logiciel metier" -NoCount

    $primary  = "Oracle.JavaRuntimeEnvironment"
    $fallback = "EclipseAdoptium.Temurin.17.JRE"

    if (Test-WingetPackageInstalled -Id $primary) {
        Write-Status -Name "Java (Oracle JRE)" -Status SKIP -Detail "deja present"
        return
    }
    if (Test-WingetPackageInstalled -Id $fallback) {
        Write-Status -Name "Java (Temurin 17)" -Status SKIP -Detail "deja present"
        return
    }

    $st = Install-WingetApp -Name "Java (Oracle JRE)" -Id $primary -Retries 1
    if ($st -eq "OK" -or $st -eq "SKIP") { return }

    if ($script:Stats.Fail -gt 0) { $script:Stats.Fail-- }
    Write-Status -Name "Java" -Status INFO -Detail ("Oracle indisponible " + $UI.Arrow + " Temurin 17 JRE") -NoCount
    Install-WingetApp -Name "Java (Temurin 17 JRE)" -Id $fallback -Retries 2 | Out-Null
}


function Save-ProvDownload {
    <# Telechargement robuste : suit les redirects 301/302/308 (curl -L), puis BITS, puis IWR. #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$OutFile,
        [string]$UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    )
    $dir = Split-Path -Parent $OutFile
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force -ErrorAction SilentlyContinue | Out-Null
    }
    if (Test-Path $OutFile) { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue }

    # 1) curl.exe (Win10/11) — suit 308 Permanent Redirect + UA navigateur (anti-403 CDN)
    $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
    if ($curl) {
        try {
            $curlArgs = @("-fsSL","--retry","3","--connect-timeout","30","-A",$UserAgent,"-o",$OutFile,$Url)
            $p = Start-Process -FilePath "curl.exe" -ArgumentList $curlArgs -Wait -PassThru -WindowStyle Hidden
            if (($p.ExitCode -eq 0) -and (Test-Path $OutFile) -and ((Get-Item $OutFile).Length -gt 1024)) {
                Write-Log "Download OK (curl) $Url -> $OutFile ($((Get-Item $OutFile).Length) o)" -Level INFO
                return $true
            }
            Write-Log "curl exit=$($p.ExitCode) pour $Url (size=$((Get-Item $OutFile -EA SilentlyContinue).Length))" -Level WARN
        } catch {
            Write-Log "curl fail: $($_.Exception.Message)" -Level WARN
        }
    }

    # 2) BITS
    try {
        Start-BitsTransfer -Source $Url -Destination $OutFile -ErrorAction Stop
        if ((Test-Path $OutFile) -and ((Get-Item $OutFile).Length -gt 1024)) {
            Write-Log "Download OK (BITS) $Url" -Level INFO
            return $true
        }
    } catch {
        Write-Log "BITS fail: $($_.Exception.Message)" -Level WARN
    }

    # 3) HttpClient .NET (suit redirects y compris 308)
    try {
        Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue
        $handler = New-Object System.Net.Http.HttpClientHandler
        $handler.AllowAutoRedirect = $true
        $client = New-Object System.Net.Http.HttpClient($handler)
        $client.Timeout = [TimeSpan]::FromMinutes(15)
        try { $client.DefaultRequestHeaders.UserAgent.ParseAdd($UserAgent) } catch {}
        $bytes = $client.GetByteArrayAsync($Url).Result
        [System.IO.File]::WriteAllBytes($OutFile, $bytes)
        if ((Test-Path $OutFile) -and ((Get-Item $OutFile).Length -gt 1024)) {
            Write-Log "Download OK (HttpClient) $Url ($($bytes.Length) o)" -Level INFO
            return $true
        }
    } catch {
        Write-Log "HttpClient fail: $($_.Exception.Message)" -Level WARN
    }

    Write-Log "Download ECHEC: $Url" -Level ERR
    return $false
}

function Test-AnyDeskInstalled {
    $paths = @(
        "${env:ProgramFiles}\AnyDesk\AnyDesk.exe",
        "${env:ProgramFiles(x86)}\AnyDesk\AnyDesk.exe"
    )
    if ($paths | Where-Object { Test-Path $_ }) { return $true }
    $u = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match 'AnyDesk' }
    return [bool]$u
}

function Install-AnyDeskDirect {
    <# Install AnyDesk hors winget (telechargement officiel + --silent). #>
    Write-Status -Name "AnyDesk" -Status RUN -Detail "install direct (hors winget)" -NoCount

    if (Test-AnyDeskInstalled) {
        Write-Status -Name "AnyDesk" -Status SKIP -Detail "deja present (detecte sur disque)"
        return "SKIP"
    }

    $localCandidates = @(
        (Find-ProvOfflineFile -Names @("AnyDesk.exe")),
        (Join-Path $ScriptDir "AnyDesk.exe"),
        (Join-Path $ScriptDir "offline\AnyDesk.exe"),
        (Join-Path $ScriptDir "ninite\AnyDesk.exe")
    ) | Where-Object { $_ } | Select-Object -Unique
    $exe = $null
    foreach ($c in $localCandidates) {
        if ($c -and (Test-Path -LiteralPath $c)) { $exe = $c; break }
    }

    if (-not $exe) {
        $exe = Join-Path $env:TEMP ("{0}-AnyDesk.exe" -f $ProvConfig.WorkFolderName)
        $urls = @(
            "https://download.anydesk.com/AnyDesk.exe",
            "https://download.anydesk.com/current/win32/AnyDesk.exe"
        )
        $okDl = $false
        foreach ($u in $urls) {
            Write-Status -Name "AnyDesk" -Status RUN -Detail "telechargement $u" -NoCount
            if (Save-ProvDownload -Url $u -OutFile $exe) { $okDl = $true; break }
        }
        if (-not $okDl) {
            Write-Status -Name "AnyDesk" -Status FAIL -Detail "telechargement impossible"
            return "FAIL"
        }
    } else {
        Write-Status -Name "AnyDesk" -Status INFO -Detail "utilise copie locale: $exe" -NoCount
    }

    $target = "${env:ProgramFiles(x86)}\AnyDesk"
    if (-not (Test-Path "${env:ProgramFiles(x86)}")) { $target = "${env:ProgramFiles}\AnyDesk" }

    Write-Status -Name "AnyDesk" -Status RUN -Detail "installation silencieuse..." -NoCount
    try {
        $args = "--install `"$target`" --silent --start-with-win --create-shortcuts"
        $p = Start-Process -FilePath $exe -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
        Write-Log "AnyDesk install exit=$($p.ExitCode)" -Level INFO
        Start-Sleep -Seconds 2
        if (Test-AnyDeskInstalled) {
            Write-Status -Name "AnyDesk" -Status OK -Detail "installe (direct)"
            return "OK"
        }
        Write-Status -Name "AnyDesk" -Status FAIL -Detail "exit $($p.ExitCode) — binaire absent"
        return "FAIL"
    } catch {
        Write-Status -Name "AnyDesk" -Status FAIL -Detail $_.Exception.Message
        return "FAIL"
    }
}

function Get-ProvOfflineDirsOnly {
    <# Uniquement les dossiers nommes 'offline' (pas usb_pro/downloads, pas ScriptDir). #>
    $out = @()
    foreach ($r in @(Get-ProvOfflineRoots)) {
        $leaf = Split-Path -Leaf $r
        if ($leaf -eq 'offline') { $out += $r }
    }
    return $out
}

function Find-ProvInOfflineFolder {
    param([string[]]$Names)
    foreach ($root in @(Get-ProvOfflineDirsOnly)) {
        foreach ($n in $Names) {
            if ([string]::IsNullOrWhiteSpace($n)) { continue }
            $p = Join-Path $root $n
            if (Test-Path -LiteralPath $p -PathType Leaf) { return $p }
        }
    }
    return $null
}

function Invoke-ProvNinite {
    <#
    Ninite UNIQUEMENT depuis offline\ (jamais usb_pro/downloads ni racine cle).
    Placez le Ninite (config NiniteFileNames) dans scripts\offline\
    #>
    $nin = Find-ProvInOfflineFolder -Names @($ProvConfig.NiniteFileNames)
    if (-not $nin) {
        foreach ($root in @(Get-ProvOfflineDirsOnly)) {
            $hit = Get-ChildItem -LiteralPath $root -File -Filter "ninite*.exe" -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if ($hit) { $nin = $hit.FullName; break }
        }
    }
    # Fallback explicite ScriptDir\offline uniquement (pas ScriptDir racine)
    if (-not $nin) {
        $localOff = Join-Path $ScriptDir "offline"
        foreach ($n in @($ProvConfig.NiniteFileNames)) {
            $p = Join-Path $localOff $n
            if (Test-Path -LiteralPath $p) { $nin = $p; break }
        }
    }

    if (-not $nin) {
        Write-Status -Name "Ninite" -Status INFO -Detail "ninite pas trouve — assurez-vous qu'il soit dans le dossier offline\" -NoCount
        return
    }

    Write-Host ""
    Write-Host ("  " + $UI.VLine + "  Ninite detecte :") -ForegroundColor Gray
    Write-Host ("  " + $UI.VLine + "  $nin") -ForegroundColor DarkGray
    $ans = (Read-Host "Lancer Ninite maintenant ? [o/N]").Trim().ToLowerInvariant()
    if ($ans -notin @('o', 'oui', 'y', 'yes')) {
        Write-Status -Name "Ninite" -Status SKIP -Detail "refuse par l'operateur"
        return
    }

    Write-Status -Name "Ninite" -Status RUN -Detail $nin -NoCount
    try {
        $p = Start-Process -FilePath $nin -Wait -PassThru
        Write-Log "Ninite exit=$($p.ExitCode)" -Level INFO
        Write-Status -Name "Ninite" -Status OK -Detail "termine (code $($p.ExitCode))"
    } catch {
        Write-Status -Name "Ninite" -Status FAIL -Detail $_.Exception.Message
    }
}


function Test-OfficeBinariesPresent {
    $bins = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE",
        "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE",
        "${env:ProgramFiles}\Microsoft Office\root\Office16\OUTLOOK.EXE"
    )
    return [bool]($bins | Where-Object { Test-Path $_ } | Select-Object -First 1)
}

function Wait-OfficeBinaries {
    param([int]$TimeoutSec = 900)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        if (Test-OfficeBinariesPresent) { return $true }
        Start-Sleep -Seconds 5
    }
    return (Test-OfficeBinariesPresent)
}

function Install-OfficeClickToRun {
    <#
    Methode officielle entreprise : Office Deployment Tool + setup.exe /configure
    Produit O365BusinessRetail (Microsoft 365 Apps for business / Business Premium), FR x64.
    Necessaire car winget Microsoft.Office est souvent en hash mismatch.
    #>
    $work = Join-Path $env:TEMP ("{0}-OfficeC2R" -f $ProvConfig.WorkFolderName)
    New-Item -ItemType Directory -Path $work -Force -ErrorAction SilentlyContinue | Out-Null
    $odtExe = Join-Path $work "officedeploymenttool.exe"
    $cfg    = Join-Path $work "configuration.xml"

    # Priorite : ODT sur cle USB / offline (puis a cote du script)
    $localOdt = Find-ProvOfflineFile -Names @("officedeploymenttool.exe", "odt.exe")
    if (-not $localOdt) {
        $localOdt = @(
            (Join-Path $ScriptDir "officedeploymenttool.exe"),
            (Join-Path $ScriptDir "offline\officedeploymenttool.exe"),
            (Join-Path $ScriptDir "offline\odt.exe")
        ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    }

    $got = $false
    if ($localOdt) {
        Copy-Item -LiteralPath $localOdt -Destination $odtExe -Force
        $got = ((Test-Path $odtExe) -and ((Get-Item $odtExe).Length -gt 500KB))
        if ($got) { Write-Status -Name "Office C2R" -Status INFO -Detail "ODT local: $localOdt" -NoCount }
    }

    if (-not $got) {
        Write-Status -Name "Office C2R" -Status RUN -Detail "telechargement ODT (curl, suit redirect 308)..." -NoCount
        $urls = @(
            "https://aka.ms/officedeploymenttool",
            "https://download.microsoft.com/download/6c44cde7-f681-471f-9635-8e6caac8dc32/officedeploymenttool_18925-20138.exe"
        )
        foreach ($url in $urls) {
            if (Save-ProvDownload -Url $url -OutFile $odtExe) {
                if ((Get-Item $odtExe).Length -gt 500KB) { $got = $true; break }
            }
        }
    }

    if (-not $got) {
        Write-Status -Name "Office C2R" -Status FAIL -Detail "ODT introuvable — placez officedeploymenttool.exe dans offline\"
        return $false
    }

    Write-Status -Name "Office C2R" -Status RUN -Detail "extraction setup.exe..." -NoCount
    try {
        $extract = Start-Process -FilePath $odtExe -ArgumentList "/quiet /extract:`"$work`"" -Wait -PassThru -WindowStyle Hidden
        Write-Log "ODT extract exit=$($extract.ExitCode)" -Level INFO
    } catch {
        Write-Status -Name "Office C2R" -Status FAIL -Detail $_.Exception.Message
        return $false
    }

    $setup = Join-Path $work "setup.exe"
    if (-not (Test-Path $setup)) {
        $hit = Get-ChildItem -Path $work -Filter "setup.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hit) { $setup = $hit.FullName }
    }
    if (-not (Test-Path $setup)) {
        Write-Status -Name "Office C2R" -Status FAIL -Detail "setup.exe introuvable apres extraction"
        return $false
    }

    @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current" MigrateArch="TRUE">
    <Product ID="O365BusinessRetail">
      <Language ID="fr-fr" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Bing" />
      <ExcludeApp ID="OneDrive" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <RemoveMSI />
</Configuration>
"@ | Set-Content -Path $cfg -Encoding UTF8

    Write-Status -Name "Office C2R" -Status RUN -Detail "installation silencieuse Business (5-15 min)..." -NoCount
    Write-Log "Lance: `"$setup`" /configure `"$cfg`"" -Level INFO
    try {
        $p = Start-Process -FilePath $setup -ArgumentList "/configure `"$cfg`"" -Wait -PassThru -WindowStyle Hidden
        Write-Log "setup /configure exit=$($p.ExitCode)" -Level INFO
    } catch {
        Write-Status -Name "Office C2R" -Status FAIL -Detail $_.Exception.Message
        return $false
    }

    if (Test-OfficeBinariesPresent) {
        Write-Status -Name "Microsoft 365 Apps" -Status OK -Detail "installe (C2R Business Premium)"
        return $true
    }

    Write-Status -Name "Office C2R" -Status RUN -Detail "attente binaires Word/Excel (jusqu'a 15 min)..." -NoCount
    if (Wait-OfficeBinaries -TimeoutSec 900) {
        Write-Status -Name "Microsoft 365 Apps" -Status OK -Detail "installe (C2R Business Premium)"
        return $true
    }

    Write-Status -Name "Microsoft 365 Apps" -Status FAIL -Detail "C2R termine sans binaires - voir journal"
    return $false
}

function Install-Microsoft365 {
    Write-Status -Name "Microsoft 365 Apps" -Status RUN -Detail ("Business Premium " + $UI.DotMid + " winget puis C2R") -NoCount

    if ((Test-WingetPackageInstalled -Id "Microsoft.Office") -or (Test-OfficeBinariesPresent)) {
        Write-Status -Name "Microsoft 365 Apps" -Status SKIP -Detail "deja present"
        return $true
    }

    $st = Install-WingetApp -Name "Microsoft 365 Apps" -Id "Microsoft.Office" -Retries 0
    if ($st -eq "OK" -or $st -eq "SKIP") { return $true }
    if (Test-OfficeBinariesPresent) {
        Write-Status -Name "Microsoft 365 Apps" -Status OK -Detail "binaires detectes"
        return $true
    }

    if ($script:Stats.Fail -gt 0) { $script:Stats.Fail-- }

    Write-Status -Name "Microsoft 365 Apps" -Status INFO -Detail "winget KO -> Click-to-Run officiel (ODT)" -NoCount
    return (Install-OfficeClickToRun)
}

function ConvertTo-ProvEmailLocalPart {
    <# Normalise une chaine pour email : minuscules, sans accents, sans espaces. #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = $Text.Trim().ToLowerInvariant()
    # Supprimer accents (formulaire D + retrait marques)
    $norm = $t.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $norm.ToCharArray()) {
        $cat = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $t = $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
    # Garder lettres/chiffres uniquement
    return (($t -replace '[^a-z0-9]', ''))
}

function Start-OfficeSignIn {
    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " Connexion Microsoft 365 " + ($UI.HLine * 27)) -ForegroundColor Magenta
    Write-Host ("  " + $UI.VLine + "  Format email : 1ere lettre prenom + nom @$($ProvConfig.EmailDomain)") -ForegroundColor Gray
    Write-Host ("  " + $UI.VLine + "  Connexion via navigateur (pas Word) — saisir le mot de passe") -ForegroundColor DarkGray
    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
    Write-Host ""

    $email = $null
    if ($script:ProvEmail) {
        Write-Host ("     Suggestion : $($script:ProvEmail)") -ForegroundColor Cyan
        $confirm = (Read-Host "     Email [Entree = suggestion, ou saisir un autre]").Trim()
        if ([string]::IsNullOrWhiteSpace($confirm)) { $email = $script:ProvEmail } else { $email = $confirm }
    } else {
        $prenom = (Read-Host "     Prenom").Trim()
        $nom    = (Read-Host "     Nom").Trim()
        $pPart = ConvertTo-ProvEmailLocalPart $prenom
        $nPart = ConvertTo-ProvEmailLocalPart $nom
        if ($pPart.Length -ge 1 -and $nPart.Length -ge 1) {
            $suggested = ($pPart.Substring(0, 1) + $nPart) + "@$($ProvConfig.EmailDomain)"
            Write-Host ("     Suggestion : $suggested") -ForegroundColor Cyan
            $confirm = (Read-Host "     Email [Entree = suggestion]").Trim()
            $email = if ([string]::IsNullOrWhiteSpace($confirm)) { $suggested } else { $confirm }
        } else {
            $email = (Read-Host "     Email professionnel").Trim()
        }
    }
    $script:ProvEmail = $email
    $prenom = $script:ProvPrenom
    $nom = $script:ProvNom

    if ([string]::IsNullOrWhiteSpace($email) -or $email -notmatch "^[^@\s]+@[^@\s]+\.[^@\s]+$") {
        Write-Status -Name "Email" -Status FAIL -Detail ("invalide " + $UI.EmDash + " connexion ignoree")
        return
    }

    Write-Status -Name "Email" -Status OK -Detail $email
    Write-Log "Email Office : $email (prenom=$prenom nom=$nom)" -Level INFO

    # Navigateur uniquement — ouvrir Word forçait une app inutile et confondait l'operateur
    $url = "https://www.office.com/login?login_hint=$([uri]::EscapeDataString($email))"
    try {
        Start-Process $url
        Write-Status -Name "Connexion M365" -Status OK -Detail "navigateur ouvert (email prerempli)"
    } catch {
        Write-Status -Name "Connexion M365" -Status FAIL -Detail $_.Exception.Message
        return
    }

    Write-Host ""
    Write-Host ("  " + $UI.VLine + "  Connectez l'utilisateur (mot de passe / MFA) dans le navigateur.") -ForegroundColor Yellow
    Write-Host ("  " + $UI.VLine + "  Le script attend votre confirmation avant de continuer.") -ForegroundColor DarkGray
    $connected = $false
    while (-not $connected) {
        $connected = Read-ProvYesNo -Prompt "Utilisateur connecte a Microsoft 365 ?" -Default Yes
        if (-not $connected) {
            Write-Host ("     Pas encore ? Terminez la connexion puis repondez O.") -ForegroundColor Yellow
            # Relance le login si besoin
            try { Start-Process $url } catch {}
        }
    }
    Write-Status -Name "Connexion M365" -Status OK -Detail "confirmee par l'operateur"
}

# ============================================================================
#  CHROME PAR DEFAUT (navigateur + PDF)
# ============================================================================

function Find-ChromeExe {
    @(
        "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Install-ChromeDirect {
    <#
      Install Chrome hors winget (comme AnyDesk) :
        1) MSI/EXE dans offline\
        2) Telechargement MSI enterprise Google (URL officielle)
      Silencieux : msiexec /qn  ou  ChromeStandaloneSetup64.exe /silent /install
    #>
    Write-Status -Name "Google Chrome" -Status RUN -Detail "install direct (hors winget)" -NoCount

    if (Find-ChromeExe) {
        Write-Status -Name "Google Chrome" -Status SKIP -Detail "deja present (chrome.exe)"
        return "SKIP"
    }

    $local = @(
        (Find-ProvOfflineFile -Names @(
            "googlechromestandaloneenterprise64.msi",
            "GoogleChromeStandaloneEnterprise64.msi",
            "ChromeStandaloneSetup64.exe",
            "ChromeSetup.exe",
            "chrome.msi"
        )),
        (Join-Path $ScriptDir "offline\googlechromestandaloneenterprise64.msi"),
        (Join-Path $ScriptDir "offline\ChromeStandaloneSetup64.exe")
    ) | Where-Object { $_ } | Select-Object -Unique

    $installer = $null
    foreach ($c in $local) {
        if ($c -and (Test-Path -LiteralPath $c)) { $installer = $c; break }
    }

    if (-not $installer) {
        $tmpDir = Join-Path $env:TEMP ("{0}-Chrome" -f $ProvConfig.WorkFolderName)
        New-Item -ItemType Directory -Path $tmpDir -Force -EA SilentlyContinue | Out-Null
        $msi = Join-Path $tmpDir "googlechromestandaloneenterprise64.msi"
        $urls = @(
            "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi",
            "https://dl.google.com/chrome/install/GoogleChromeStandaloneEnterprise64.msi"
        )
        $okDl = $false
        foreach ($u in $urls) {
            Write-Status -Name "Google Chrome" -Status RUN -Detail "telechargement MSI Google..." -NoCount
            if (Save-ProvDownload -Url $u -OutFile $msi) {
                if ((Test-Path $msi) -and ((Get-Item $msi).Length -gt 1MB)) {
                    $installer = $msi
                    $okDl = $true
                    break
                }
            }
        }
        if (-not $okDl) {
            # Fallback EXE offline installer
            $exe = Join-Path $tmpDir "ChromeStandaloneSetup64.exe"
            $exeUrls = @(
                "https://dl.google.com/chrome/install/ChromeStandaloneSetup64.exe",
                "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7B%7D%26lang%3Dfr%26browser%3D0%26usagestats%3D0%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers/update2/installers/ChromeStandaloneSetup64.exe"
            )
            foreach ($u in $exeUrls) {
                Write-Status -Name "Google Chrome" -Status RUN -Detail "telechargement EXE Google..." -NoCount
                if (Save-ProvDownload -Url $u -OutFile $exe) {
                    if ((Test-Path $exe) -and ((Get-Item $exe).Length -gt 1MB)) {
                        $installer = $exe
                        $okDl = $true
                        break
                    }
                }
            }
        }
        if (-not $okDl -or -not $installer) {
            Write-Status -Name "Google Chrome" -Status FAIL -Detail "telechargement impossible — mettre le MSI dans offline\"
            return "FAIL"
        }
    } else {
        Write-Status -Name "Google Chrome" -Status INFO -Detail "utilise copie locale: $installer" -NoCount
    }

    Write-Status -Name "Google Chrome" -Status RUN -Detail "installation silencieuse..." -NoCount
    try {
        $ext = [IO.Path]::GetExtension($installer).ToLowerInvariant()
        if ($ext -eq ".msi") {
            $p = Start-Process -FilePath "msiexec.exe" -ArgumentList @(
                "/i", "`"$installer`"", "/qn", "/norestart", "ALLUSERS=1"
            ) -Wait -PassThru -WindowStyle Hidden
        } else {
            $p = Start-Process -FilePath $installer -ArgumentList "/silent /install" -Wait -PassThru -WindowStyle Hidden
        }
        Write-Log "Chrome direct install exit=$($p.ExitCode) file=$installer" -Level INFO
        Start-Sleep -Seconds 3
        if (Find-ChromeExe) {
            Write-Status -Name "Google Chrome" -Status OK -Detail "installe (direct MSI/EXE)"
            return "OK"
        }
        # msiexec 0 / 3010 = OK ; parfois chrome.exe apparait apres un delai
        if ($p.ExitCode -in @(0, 3010)) {
            Start-Sleep -Seconds 5
            if (Find-ChromeExe) {
                Write-Status -Name "Google Chrome" -Status OK -Detail "installe (direct, delai)"
                return "OK"
            }
        }
        Write-Status -Name "Google Chrome" -Status FAIL -Detail ("exit $($p.ExitCode) — chrome.exe absent")
        return "FAIL"
    } catch {
        Write-Status -Name "Google Chrome" -Status FAIL -Detail $_.Exception.Message
        return "FAIL"
    }
}

function Install-ChromeWithFallback {
    <# winget Google.Chrome → (opt) Google.Chrome.EXE → MSI/EXE direct Google / offline. #>
    if (Find-ChromeExe) {
        Write-Status -Name "Google Chrome" -Status SKIP -Detail "deja present (chrome.exe)"
        return "SKIP"
    }

    $st = Install-WingetApp -Name "Google Chrome" -Id "Google.Chrome" -Retries 1
    if (Find-ChromeExe) {
        if ($st -eq "FAIL" -and $script:Stats.Fail -gt 0) { $script:Stats.Fail-- }
        if ($st -eq "FAIL") {
            Write-Status -Name "Google Chrome" -Status OK -Detail "present apres winget (code ignore)"
            return "OK"
        }
        return $st
    }

    # Package EXE winget (URLs versionnees — parfois OK quand le MSI vanity casse)
    if ($st -eq "FAIL" -and $script:Stats.Fail -gt 0) { $script:Stats.Fail-- }
    Write-Status -Name "Google Chrome" -Status INFO -Detail "winget MSI KO → essai Google.Chrome.EXE" -NoCount
    $stExe = Install-WingetApp -Name "Google Chrome (EXE)" -Id "Google.Chrome.EXE" -Retries 0
    if (Find-ChromeExe) {
        if ($stExe -eq "FAIL" -and $script:Stats.Fail -gt 0) { $script:Stats.Fail-- }
        return "OK"
    }

    if ($stExe -eq "FAIL" -and $script:Stats.Fail -gt 0) { $script:Stats.Fail-- }
    Write-Status -Name "Google Chrome" -Status INFO -Detail "winget KO → install direct Google" -NoCount
    return (Install-ChromeDirect)
}

function Set-ProvFirefoxDefaultBrowserOff {
    <# Empeche Firefox de se remettre navigateur par defaut / agent fond. #>
    # policies.json (Enterprise)
    $ffRoots = @(
        "${env:ProgramFiles}\Mozilla Firefox",
        "${env:ProgramFiles(x86)}\Mozilla Firefox"
    )
    $policyJson = @'
{
  "policies": {
    "DontCheckDefaultBrowser": true,
    "OverrideFirstRunPage": "",
    "OverridePostUpdatePage": ""
  }
}
'@
    foreach ($root in $ffRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        try {
            $dist = Join-Path $root "distribution"
            New-Item -ItemType Directory -Path $dist -Force -EA SilentlyContinue | Out-Null
            $utf8 = New-Object System.Text.UTF8Encoding $false
            [System.IO.File]::WriteAllText((Join-Path $dist "policies.json"), $policyJson.Trim(), $utf8)
            Write-Log "Firefox policies.json ecrit dans $dist" -Level INFO
        } catch {
            Write-Log "Firefox policies : $($_.Exception.Message)" -Level WARN
        }
    }
    # Agent "Default Browser" (tache planifiee Mozilla)
    try {
        Get-ScheduledTask -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskName -match '(?i)firefox|mozilla' -and $_.TaskName -match '(?i)default|browser|update' } |
            ForEach-Object {
                try {
                    Disable-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -EA Stop | Out-Null
                    Write-Log "Firefox task OFF: $($_.TaskPath)$($_.TaskName)" -Level INFO
                } catch {}
            }
    } catch {}
}

function Test-ProvHttpDefaultIsChrome {
    <# Best-effort : ProgId UserChoice http/https contient Chrome ? #>
    $ok = $false
    $sid = Get-InteractiveUserSid
    $roots = @("HKCU:")
    if ($sid) { $roots += "Registry::HKEY_USERS\$sid" }
    foreach ($root in $roots) {
        foreach ($scheme in @('http', 'https')) {
            $p = Join-Path $root "Software\Microsoft\Windows\Shell\Associations\UrlAssociations\$scheme\UserChoice"
            try {
                if (-not (Test-Path -LiteralPath $p)) { continue }
                $prog = [string](Get-ItemProperty -LiteralPath $p -EA Stop).ProgId
                Write-Log "UserChoice $scheme ($root) = $prog" -Level INFO
                if ($prog -match '(?i)chrome') { $ok = $true }
            } catch {}
        }
    }
    return $ok
}

function Set-ChromeAsDefault {
    Write-Status -Name "Chrome par defaut" -Status RUN -Detail "http / https / html / pdf" -NoCount

    $chrome = Find-ChromeExe
    if (-not $chrome) {
        Write-Status -Name "Chrome par defaut" -Status FAIL -Detail "chrome.exe introuvable"
        return
    }

    try {
        $pol = "HKLM:\SOFTWARE\Policies\Google\Chrome"
        if (-not (Test-Path $pol)) { New-Item -Path $pol -Force | Out-Null }
        New-ItemProperty -Path $pol -Name "DefaultBrowserSettingEnabled" -PropertyType DWord -Value 1 -Force -ErrorAction SilentlyContinue | Out-Null
    } catch {
        Write-Log "Policy Chrome : $($_.Exception.Message)" -Level WARN
    }

    # ProgId machine (utile pour DISM / ftype)
    try {
        $cmd = "`"$chrome`" -- `"%1`""
        New-Item -Path "HKLM:\SOFTWARE\Classes\ChromeHTML\shell\open\command" -Force -EA SilentlyContinue | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\ChromeHTML\shell\open\command" -Name "(default)" -Value $cmd -Force -EA SilentlyContinue
    } catch {}

    Set-ProvFirefoxDefaultBrowserOff

    $assocXml = Join-Path $env:TEMP ("{0}-default-associations.xml" -f $ProvConfig.FilePrefix)
    @"
<?xml version="1.0" encoding="UTF-8"?>
<DefaultAssociations>
  <Association Identifier=".htm"   ProgId="ChromeHTML" ApplicationName="Google Chrome" />
  <Association Identifier=".html"  ProgId="ChromeHTML" ApplicationName="Google Chrome" />
  <Association Identifier=".shtml" ProgId="ChromeHTML" ApplicationName="Google Chrome" />
  <Association Identifier=".webp"  ProgId="ChromeHTML" ApplicationName="Google Chrome" />
  <Association Identifier=".xhtml" ProgId="ChromeHTML" ApplicationName="Google Chrome" />
  <Association Identifier=".svg"   ProgId="ChromeHTML" ApplicationName="Google Chrome" />
  <Association Identifier=".pdf"   ProgId="ChromeHTML" ApplicationName="Google Chrome" />
  <Association Identifier="http"   ProgId="ChromeHTML" ApplicationName="Google Chrome" />
  <Association Identifier="https"  ProgId="ChromeHTML" ApplicationName="Google Chrome" />
</DefaultAssociations>
"@ | Set-Content -Path $assocXml -Encoding UTF8

    $dismOk = $false
    try {
        $dism = Start-Process dism.exe -ArgumentList "/Online","/Import-DefaultAppAssociations:`"$assocXml`"" `
            -Wait -PassThru -WindowStyle Hidden
        $dismOk = ($dism.ExitCode -eq 0)
        Write-Log "DISM associations exit=$($dism.ExitCode)" -Level INFO
    } catch {
        Write-Log "DISM : $($_.Exception.Message)" -Level WARN
    }

    try {
        Start-Process -FilePath $chrome -ArgumentList "--make-default-browser" -WindowStyle Hidden
        Start-Sleep -Seconds 2
    } catch {
        Write-Log "chrome --make-default-browser : $($_.Exception.Message)" -Level WARN
    }

    try {
        cmd /c "ftype ChromeHTML=`"$chrome`" -- `"%1`"" | Out-Null
        foreach ($ext in @(".html",".htm",".pdf",".shtml",".xhtml",".webp")) {
            cmd /c "assoc $ext=ChromeHTML" | Out-Null
        }
    } catch {
        Write-Log "assoc/ftype : $($_.Exception.Message)" -Level WARN
    }

    # Win11 protege UserChoice (hash) : DISM couvre surtout les *nouveaux* profils.
    # Si le user courant n'a pas Chrome → ouvrir Parametres + OK tech.
    $userOk = Test-ProvHttpDefaultIsChrome
    if (-not $userOk) {
        Write-Host ("  " + $UI.VLine + "  Win11 : choisir Google Chrome comme navigateur par defaut") -ForegroundColor Yellow
        Write-Host ("  " + $UI.VLine + "  Parametres → Applications → Applications par defaut → Google Chrome → Definir par defaut") -ForegroundColor DarkGray
        try { Start-Process "ms-settings:defaultapps" } catch {}
        try { Start-Process -FilePath $chrome -ArgumentList "--make-default-browser" } catch {}
        $null = Read-ProvYesNo -Prompt "Chrome est bien le navigateur (et PDF) par defaut ?" -Default Yes
        $userOk = Test-ProvHttpDefaultIsChrome
    }

    if ($dismOk -or $userOk) {
        Write-Status -Name "Chrome par defaut" -Status OK -Detail $(if ($userOk) { "web + PDF (user)" } else { "DISM OK (nouveaux profils) + prompt" })
    } else {
        Write-Status -Name "Chrome par defaut" -Status OK -Detail ("applique " + $UI.EmDash + " verifier Settings si besoin")
    }
}

# ============================================================================
#  BARRE DES TACHES + OPTIMS WINDOWS
# ============================================================================

function Get-InteractiveUserSid {
    try {
        $sessionId = (Get-Process -Id $PID).SessionId
        $explorer = Get-Process explorer -ErrorAction SilentlyContinue |
            Where-Object { $_.SessionId -eq $sessionId } |
            Select-Object -First 1
        if (-not $explorer) {
            $explorer = Get-Process explorer -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        if ($explorer) {
            $proc = Get-CimInstance Win32_Process -Filter "ProcessId=$($explorer.Id)" -ErrorAction Stop
            $owner = Invoke-CimMethod -InputObject $proc -MethodName GetOwner -ErrorAction Stop
            $acct = New-Object System.Security.Principal.NTAccount($owner.Domain, $owner.User)
            return $acct.Translate([System.Security.Principal.SecurityIdentifier]).Value
        }
    } catch {}
    try {
        $user = (Get-CimInstance Win32_ComputerSystem).UserName
        if ($user) {
            $acct = New-Object System.Security.Principal.NTAccount($user)
            return $acct.Translate([System.Security.Principal.SecurityIdentifier]).Value
        }
    } catch {}
    return $null
}

function Set-RegistryDword {
    <# Ecriture registre silencieuse (pas de spam rouge). Gere HKCU eleve via HKU\SID. #>
    param([string]$Path, [string]$Name, [int]$Value)

    function Write-OneKey([string]$PsPath) {
        try {
            if (-not (Test-Path -LiteralPath $PsPath)) {
                New-Item -Path $PsPath -Force -ErrorAction Stop | Out-Null
            }
            Set-ItemProperty -LiteralPath $PsPath -Name $Name -Value $Value -Type DWord -Force -ErrorAction Stop
            return $true
        } catch {
            return $false
        }
    }

    $ok = Write-OneKey $Path

    # Fallback reg.exe (souvent OK quand Set-ItemProperty est refuse)
    if (-not $ok) {
        $winPath = $Path -replace '^HKCU:', 'HKCU' -replace '^HKLM:', 'HKLM'
        $winPath = $winPath -replace '\\', '\'
        # Normaliser HKCU\Software\...
        if ($winPath -match '^HKCU\\') { }
        elseif ($Path -like 'HKCU:*') { $winPath = 'HKCU' + ($Path.Substring(5)) }
        elseif ($Path -like 'HKLM:*') { $winPath = 'HKLM' + ($Path.Substring(5)) }
        $winPath = $winPath -replace '/', '\'
        & reg.exe add "$winPath" /v "$Name" /t REG_DWORD /d $Value /f 1>$null 2>$null
        if ($LASTEXITCODE -eq 0) { $ok = $true }
    }

    # Si HKCU et session elevee : ecrire aussi dans la ruche de l'utilisateur interactif
    if ($Path -like 'HKCU:*') {
        $sid = Get-InteractiveUserSid
        if ($sid) {
            $rel = $Path.Substring(5)  # apres HKCU:
            if (-not $rel.StartsWith('\')) { $rel = '\' + $rel }
            $hkuPs = "Registry::HKEY_USERS\$sid$rel"
            if (Write-OneKey $hkuPs) { $ok = $true }
            $hkuWin = "HKU\$sid" + ($rel -replace '/', '\')
            & reg.exe add "$hkuWin" /v "$Name" /t REG_DWORD /d $Value /f 1>$null 2>$null
            if ($LASTEXITCODE -eq 0) { $ok = $true }
        }
    }

    if (-not $ok) {
        Write-Log "Reg KO: $Path\$Name" -Level WARN
    }
    return $ok
}

function Disable-ProvTaskbarWidgets {
    <# Enleve la meteo / Widgets de la barre des taches (Win11). #>
    Write-Status -Name "Meteo / Widgets" -Status RUN -Detail "desactivation barre des taches" -NoCount

    # Policies machine (fiables sous session admin / Intune)
    $pol1 = Set-RegistryDword "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" "AllowNewsAndInterests" 0
    $pol2 = Set-RegistryDword "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds" "EnableFeeds" 0

    # HKCU : best-effort (peut etre refuse en session elevee — plus de spam rouge)
    Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "TaskbarDa" 0 | Out-Null
    Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\Feeds" "ShellFeedsTaskbarViewMode" 2 | Out-Null
    Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowCopilotButton" 0 | Out-Null
    Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "TaskbarCopilot" 0 | Out-Null

    if ($pol1 -or $pol2) {
        Write-Status -Name "Meteo / Widgets" -Status OK -Detail "desactives (policy + HKCU)"
    } else {
        Write-Status -Name "Meteo / Widgets" -Status OK -Detail "tente via HKCU (redemarrer session si visible)"
    }
}

function Set-ProvPerformanceTweaks {
    <# Optimisations legeres pour mini PC Win11 (safe Intune / M365). #>
    Write-Status -Name "Optimisations" -Status RUN -Detail "perfs mini PC" -NoCount
    $n = 0

    # Transparence / effets visuels
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" "EnableTransparency" 0) { $n++ }
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" "VisualFXSetting" 2) { $n++ }  # 2 = best performance
    if (Set-RegistryDword "HKCU:\Control Panel\Desktop\WindowMetrics" "MinAnimate" 0) { $n++ }
    # Animations fenetre
    try {
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "MenuShowDelay" -Value "0" -Force
        $n++
    } catch {}

    # Suggestions / pubs / tips Windows
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338389Enabled" 0) { $n++ }
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SystemPaneSuggestionsEnabled" 0) { $n++ }
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\UserProfileEngagement" "ScoobeSystemSettingEnabled" 0) { $n++ }
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowSyncProviderNotifications" 0) { $n++ }
    # Bing / cloud search OFF, mais suggestions locales (apps/fichiers recents) ON
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings" "IsAADCloudSearchEnabled" 0) { $n++ }
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" "BingSearchEnabled" 0) { $n++ }
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" "CortanaConsent" 0) { $n++ }
    # Retirer la policy qui vide la loupe (ne pas laisser DisableSearchBoxSuggestions=1)
    try {
        $pol = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
        if (Test-Path $pol) {
            Remove-ItemProperty -LiteralPath $pol -Name "DisableSearchBoxSuggestions" -Force -ErrorAction SilentlyContinue
            $n++
        }
    } catch {}
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings" "IsDeviceSearchHistoryEnabled" 1) { $n++ }
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings" "IsDynamicSearchBoxEnabled" 1) { $n++ }

    # Xbox Game Bar / Game DVR (consomme GPU/CPU en fond)
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR" "AppCaptureEnabled" 0) { $n++ }
    if (Set-RegistryDword "HKCU:\System\GameConfigStore" "GameDVR_Enabled" 0) { $n++ }
    if (Set-RegistryDword "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR" "AllowGameDVR" 0) { $n++ }

    # Delai demarrage apps (Explorer)
    if (Set-RegistryDword "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize" "StartupDelayInMSec" 0) { $n++ }

    # Plan d'alimentation Haute performance
    try {
        # GUID standard "High performance"
        powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c 2>$null
        if ($LASTEXITCODE -ne 0) {
            # Fallback Ultimate Performance (si present)
            powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61 2>$null | Out-Null
            powercfg -setactive e9a42b02-d5df-448d-aa00-03f14749eb61 2>$null
        }
        $n++
        Write-Log "powercfg High Performance tente" -Level INFO
    } catch {
        Write-Log "powercfg : $($_.Exception.Message)" -Level WARN
    }

    # Telemetrie DiagTrack (leger gain + moins de bruit reseau)
    try {
        $diag = Get-Service -Name "DiagTrack" -ErrorAction SilentlyContinue
        if ($diag -and $diag.Status -ne "Stopped") {
            Stop-Service -Name "DiagTrack" -Force -ErrorAction SilentlyContinue
            Set-Service -Name "DiagTrack" -StartupType Disabled -ErrorAction SilentlyContinue
            $n++
        }
    } catch {}

    Write-Status -Name "Optimisations" -Status OK -Detail "$n reglages appliques"
}

function Test-ProvStartupDenied {
    <# Denylist explicite (prioritaire) : Firefox / Mozilla ne doivent jamais demarrer. #>
    param($Name, $Command)
    $blob = ("{0} {1}" -f ([string]$Name), (ConvertTo-ProvStartupCommandString $Command))
    foreach ($pat in @('(?i)firefox', '(?i)mozilla', '(?i)\\firefox\.exe\b')) {
        if ($blob -match $pat) { return $true }
    }
    return $false
}

function ConvertTo-ProvStartupCommandString {
    <# Evite "types ne correspondent pas" si la valeur registre est string[] / byte[]. #>
    param($Value)
    try {
        if ($null -eq $Value) { return "" }
        if ($Value -is [string]) { return $Value }
        if ($Value -is [System.Array]) {
            $parts = foreach ($x in @($Value)) { [string]$x }
            return ($parts -join ' ')
        }
        return [string]$Value
    } catch {
        return ""
    }
}

function Test-ProvStartupAllowed {
    <# Allowlist demarrage : Java, AnyDesk, Discord, Chrome, antivirus Windows. #>
    param($Name, $Command)
    if (Test-ProvStartupDenied -Name $Name -Command $Command) { return $false }
    $blob = ("{0} {1}" -f ([string]$Name), (ConvertTo-ProvStartupCommandString $Command))
    $allow = @(
        '(?i)\bjava(w)?\b|\bjre\b|\bjdk\b|adoptium|temurin|eclipse.?adoptium|oracle\\java',
        '(?i)anydesk',
        '(?i)discord',
        '(?i)chrome|google\\s*chrome|google\\chrome',
        '(?i)securityhealth|windows.?defender|msascui|defender|antispyware|smartscreen|msmpeng|securityhealthsystray'
    )
    foreach ($pat in $allow) {
        if ($blob -match $pat) { return $true }
    }
    return $false
}

function New-ProvStartupApprovedBytes {
    param([ValidateSet(2,3)][int]$Flag)
    # 02 = active, 03 = desactive — toujours byte[12] via New-Object (pas [byte[]](...))
    $b = New-Object byte[] 12
    $b[0] = [byte]$Flag
    try {
        $ft = [BitConverter]::GetBytes([int64]([DateTime]::UtcNow.ToFileTimeUtc()))
        if ($ft -and $ft.Length -ge 8) {
            [Buffer]::BlockCopy($ft, 0, $b, 4, 8)
        }
    } catch {}
    return ,$b  # comma : empeche l'unroll pipeline PS
}

function Disable-ProvStartupApprovedValue {
    param([string]$ApprovedPath, [string]$ValueName)
    if ([string]::IsNullOrWhiteSpace($ApprovedPath) -or [string]::IsNullOrWhiteSpace($ValueName)) { return $false }
    $disabled = New-ProvStartupApprovedBytes -Flag 3
    try {
        if (-not (Test-Path -LiteralPath $ApprovedPath)) {
            New-Item -Path $ApprovedPath -Force -EA Stop | Out-Null
        }
        New-ItemProperty -LiteralPath $ApprovedPath -Name $ValueName -PropertyType Binary -Value $disabled -Force -EA Stop | Out-Null
        return $true
    } catch {
        try {
            Set-ItemProperty -LiteralPath $ApprovedPath -Name $ValueName -Value $disabled -Type Binary -Force -EA Stop
            return $true
        } catch {
            return $false
        }
    }
}

function Enable-ProvStartupApprovedValue {
    param([string]$ApprovedPath, [string]$ValueName)
    if ([string]::IsNullOrWhiteSpace($ApprovedPath) -or [string]::IsNullOrWhiteSpace($ValueName)) { return $false }
    $enabled = New-ProvStartupApprovedBytes -Flag 2
    try {
        if (-not (Test-Path -LiteralPath $ApprovedPath)) {
            New-Item -Path $ApprovedPath -Force -EA Stop | Out-Null
        }
        New-ItemProperty -LiteralPath $ApprovedPath -Name $ValueName -PropertyType Binary -Value $enabled -Force -EA Stop | Out-Null
        return $true
    } catch {
        try {
            Set-ItemProperty -LiteralPath $ApprovedPath -Name $ValueName -Value $enabled -Type Binary -Force -EA Stop
            return $true
        } catch {
            return $false
        }
    }
}

function Get-ProvInteractiveRunKeySets {
    <# HKCU + HKU\SID (session elevee) + HKLM. #>
    $sets = New-Object System.Collections.Generic.List[object]
    $sid = $null
    try { $sid = Get-InteractiveUserSid } catch { $sid = $null }

    [void]$sets.Add([ordered]@{
        Path     = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
        Approved = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
    })
    if ($sid) {
        [void]$sets.Add([ordered]@{
            Path     = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Run"
            Approved = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
        })
    }
    [void]$sets.Add([ordered]@{
        Path     = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
        Approved = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
    })
    [void]$sets.Add([ordered]@{
        Path     = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"
        Approved = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32"
    })
    return [object[]]$sets.ToArray()
}

function Get-ProvInteractiveStartupFolders {
    $dirs = New-Object System.Collections.Generic.List[string]
    $sid = $null
    try { $sid = Get-InteractiveUserSid } catch { $sid = $null }
    if ($sid) {
        try {
            $ush = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            $sf  = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            foreach ($p in @($ush, $sf)) {
                if (-not (Test-Path -LiteralPath $p)) { continue }
                $v = $null
                try { $v = [string](Get-ItemProperty -LiteralPath $p -EA SilentlyContinue).Startup } catch {}
                if ([string]::IsNullOrWhiteSpace($v)) { continue }
                $prof = $null
                try {
                    $prof = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -EA Stop).ProfileImagePath
                } catch {}
                if ($prof) {
                    $v = $v -replace '(?i)%USERPROFILE%', [string]$prof
                    $v = $v -replace '(?i)%SystemDrive%', [string]$env:SystemDrive
                }
                try { $v = [Environment]::ExpandEnvironmentVariables($v) } catch {}
                if ($v -and (Test-Path -LiteralPath $v)) {
                    if (-not $dirs.Contains($v)) { [void]$dirs.Add($v) }
                }
            }
        } catch {}
    }
    try {
        $fallback = [Environment]::GetFolderPath('Startup')
        if ($fallback -and -not $dirs.Contains($fallback)) { [void]$dirs.Add($fallback) }
    } catch {}
    $pd = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
    if (-not $dirs.Contains($pd)) { [void]$dirs.Add($pd) }
    return [string[]]$dirs.ToArray()
}

function Ensure-ProvChromeStartup {
    <# Ajoute Google Chrome au demarrage (HKCU + user interactif) et l'active. #>
    param($Report = $null)
    $chrome = $null
    try { $chrome = Find-ChromeExe } catch {}
    if (-not $chrome) {
        Write-Log "Ensure Chrome startup: chrome.exe introuvable" -Level WARN
        return $false
    }
    $valueName = "Google Chrome"
    $cmd = '"' + $chrome + '"'
    $okAny = $false
    $targets = New-Object System.Collections.Generic.List[object]
    [void]$targets.Add([ordered]@{
        Path     = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
        Approved = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
    })
    $sid = $null
    try { $sid = Get-InteractiveUserSid } catch {}
    if ($sid) {
        [void]$targets.Add([ordered]@{
            Path     = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Run"
            Approved = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
        })
    }
    foreach ($t in $targets) {
        try {
            if (-not (Test-Path -LiteralPath $t.Path)) {
                New-Item -Path $t.Path -Force -EA Stop | Out-Null
            }
            New-ItemProperty -LiteralPath $t.Path -Name $valueName -PropertyType String -Value $cmd -Force -EA Stop | Out-Null
            [void](Enable-ProvStartupApprovedValue -ApprovedPath $t.Approved -ValueName $valueName)
            $okAny = $true
            Write-Log "Startup ENSURE Chrome: $($t.Path) => $cmd" -Level INFO
        } catch {
            # Fallback reg.exe (souvent OK quand New-ItemProperty type-mismatch)
            try {
                $winPath = $t.Path -replace '^HKCU:', 'HKCU' -replace '^Registry::HKEY_USERS\\', 'HKU\' -replace '^HKLM:', 'HKLM'
                & reg.exe add "$winPath" /v "$valueName" /t REG_SZ /d $cmd /f 1>$null 2>$null
                if ($LASTEXITCODE -eq 0) {
                    [void](Enable-ProvStartupApprovedValue -ApprovedPath $t.Approved -ValueName $valueName)
                    $okAny = $true
                    Write-Log "Startup ENSURE Chrome via reg.exe: $winPath" -Level INFO
                } else {
                    Write-Log "Startup ENSURE Chrome FAIL $($t.Path): $($_.Exception.Message)" -Level WARN
                }
            } catch {
                Write-Log "Startup ENSURE Chrome FAIL $($t.Path): $($_.Exception.Message)" -Level WARN
            }
        }
    }
    if ($okAny -and $null -ne $Report) {
        try { [void]$Report.Add("ON Chrome startup") } catch {}
    }
    return $okAny
}

function Set-ProvStartupApps {
    <#
    Nettoie les apps au demarrage.
    Conserve : Java, AnyDesk, Discord, Chrome, antivirus Windows.
    Force OFF : Firefox / Mozilla (tous noms).
    Force ON  : Google Chrome.
    #>
    Write-Status -Name "Demarrage Windows" -Status RUN -Detail "v3 · Chrome ON · Firefox OFF" -NoCount

    $disabled = 0
    $kept = 0
    $report = New-Object System.Collections.Generic.List[string]
    $step = "init"

    try {
        $step = "runKeys"
        $runKeys = Get-ProvInteractiveRunKeySets
        if ($null -eq $runKeys) { $runKeys = @() }
        elseif ($runKeys -is [System.Collections.IEnumerable] -and $runKeys -isnot [string] -and $runKeys -isnot [hashtable] -and $runKeys -isnot [System.Collections.IDictionary]) {
            $runKeys = @($runKeys)
        } else {
            $runKeys = @($runKeys)
        }
        # Garde-fou : chaque entree doit avoir .Path string (pas un tableau imbrique)
        $runKeys = @(
            foreach ($rk in $runKeys) {
                if ($rk -and ($rk -is [hashtable] -or $rk -is [System.Collections.IDictionary] -or $rk.PSObject.Properties['Path'])) {
                    if ($rk.Path -is [string]) { $rk }
                }
            }
        )

        $step = "scanRun"
        foreach ($rk in $runKeys) {
            if (-not $rk -or -not $rk.Path) { continue }
            if (-not (Test-Path -LiteralPath $rk.Path)) { continue }
            $props = $null
            try { $props = Get-ItemProperty -LiteralPath $rk.Path -EA Stop } catch { continue }
            if (-not $props) { continue }

            foreach ($p in @($props.PSObject.Properties)) {
                try {
                    if ($p.Name -in @('PSPath','PSParentPath','PSChildName','PSDrive','PSProvider')) { continue }
                    $name = [string]$p.Name
                    if ([string]::IsNullOrWhiteSpace($name)) { continue }
                    $cmd = ConvertTo-ProvStartupCommandString $p.Value

                    if (Test-ProvStartupAllowed -Name $name -Command $cmd) {
                        $kept++
                        [void](Enable-ProvStartupApprovedValue -ApprovedPath ([string]$rk.Approved) -ValueName $name)
                        Write-Log "Startup KEEP: $name => $cmd" -Level INFO
                        continue
                    }

                    $okDis = Disable-ProvStartupApprovedValue -ApprovedPath ([string]$rk.Approved) -ValueName $name
                    try {
                        Remove-ItemProperty -LiteralPath $rk.Path -Name $name -Force -EA Stop
                        $disabled++
                        [void]$report.Add("OFF Run: $name")
                        Write-Log "Startup DISABLE+REMOVE: $name => $cmd (approved=$okDis)" -Level INFO
                    } catch {
                        if ($okDis) {
                            $disabled++
                            [void]$report.Add("OFF Approved: $name")
                        }
                    }
                } catch {
                    Write-Log "Startup scanRun item KO: $($_.Exception.Message)" -Level WARN
                }
            }
        }

        $step = "forceOff"
        $forceOffNames = @(
            'com.squirrel.Teams.Teams', 'Teams', 'Microsoft Teams', 'TeamsMachineInstaller',
            'MicrosoftEdge', 'Microsoft Edge', 'msedge', 'Edge',
            'OneDrive', 'OneDriveSetup', 'Discord Update',
            'Spotify', 'Steam', 'iTunesHelper',
            'Mozilla Firefox', 'Firefox'
        )
        foreach ($rk in $runKeys) {
            if (-not $rk -or -not (Test-Path -LiteralPath $rk.Path)) { continue }
            try {
                $item = Get-Item -LiteralPath $rk.Path -EA SilentlyContinue
                if (-not $item) { continue }
                foreach ($vn in @($item.GetValueNames())) {
                    $hit = $false
                    foreach ($pat in $forceOffNames) {
                        if ($vn -like $pat) { $hit = $true; break }
                    }
                    if (-not $hit -and ($vn -match '(?i)firefox|mozilla')) { $hit = $true }
                    if (-not $hit) { continue }
                    $cmdVn = ConvertTo-ProvStartupCommandString ($item.GetValue($vn))
                    if (Test-ProvStartupAllowed -Name $vn -Command $cmdVn) { continue }
                    [void](Disable-ProvStartupApprovedValue -ApprovedPath ([string]$rk.Approved) -ValueName $vn)
                    Remove-ItemProperty -LiteralPath $rk.Path -Name $vn -Force -EA SilentlyContinue
                    $disabled++
                    [void]$report.Add("OFF force: $vn")
                }
            } catch {}
            try {
                if ($rk.Approved -and (Test-Path -LiteralPath $rk.Approved)) {
                    $ap = Get-Item -LiteralPath $rk.Approved -EA SilentlyContinue
                    if ($ap) {
                        foreach ($vn in @($ap.GetValueNames())) {
                            if ($vn -match '(?i)firefox|mozilla') {
                                [void](Disable-ProvStartupApprovedValue -ApprovedPath ([string]$rk.Approved) -ValueName $vn)
                                $disabled++
                                [void]$report.Add("OFF approved-ff: $vn")
                            }
                        }
                    }
                }
            } catch {}
        }

        $step = "startupFolders"
        $startupDirs = @(Get-ProvInteractiveStartupFolders)
        $backupRoot = Join-Path $env:LOCALAPPDATA (Join-Path $ProvConfig.WorkFolderName "startup-disabled")
        try { New-Item -ItemType Directory -Path $backupRoot -Force -EA SilentlyContinue | Out-Null } catch {}

        foreach ($dir in $startupDirs) {
            if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { continue }
            foreach ($lnk in @(Get-ChildItem -LiteralPath $dir -Filter "*.lnk" -Force -EA SilentlyContinue)) {
                try {
                    $target = ""
                    try {
                        $sh = New-Object -ComObject WScript.Shell
                        $target = [string]$sh.CreateShortcut($lnk.FullName).TargetPath
                    } catch {}
                    if (Test-ProvStartupAllowed -Name $lnk.BaseName -Command ("{0} {1}" -f $target, $lnk.Name)) {
                        $kept++
                        continue
                    }
                    $dest = Join-Path $backupRoot $lnk.Name
                    if (Test-Path -LiteralPath $dest) {
                        $dest = Join-Path $backupRoot ("{0}_{1}" -f ([guid]::NewGuid().ToString('N').Substring(0,6)), $lnk.Name)
                    }
                    Move-Item -LiteralPath $lnk.FullName -Destination $dest -Force -EA Stop
                    $disabled++
                    [void]$report.Add("OFF lnk: $($lnk.Name)")
                } catch {
                    try {
                        Remove-Item -LiteralPath $lnk.FullName -Force -EA Stop
                        $disabled++
                        [void]$report.Add("OFF lnk del: $($lnk.Name)")
                    } catch {}
                }
            }
        }

        $step = "ensureChrome"
        try {
            if (Ensure-ProvChromeStartup -Report $report) { $kept++ }
        } catch {
            Write-Log "Ensure Chrome startup exception: $($_.Exception.Message)" -Level WARN
            try { [void]$report.Add("WARN Chrome startup") } catch {}
        }

        $step = "firefoxPolicies"
        try { Set-ProvFirefoxDefaultBrowserOff } catch {
            Write-Log "Firefox policies KO: $($_.Exception.Message)" -Level WARN
        }

        $step = "scheduledTasks"
        try {
            Get-ScheduledTask -EA SilentlyContinue |
                Where-Object {
                    (
                        $_.TaskName -match '(?i)teams|onedrive|edgeupdate|skype' -or
                        ($_.TaskName -match '(?i)firefox|mozilla' -and $_.TaskName -match '(?i)default|browser')
                    ) -and
                    $_.TaskName -notmatch '(?i)defender|security|GoogleUpdate' -and
                    $_.State -ne 'Disabled'
                } |
                ForEach-Object {
                    try {
                        Disable-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -EA Stop | Out-Null
                        $disabled++
                        [void]$report.Add("OFF task: $($_.TaskName)")
                    } catch {}
                }
        } catch {}

        $step = "teams"
        try {
            [void](Set-RegistryDword "HKCU:\Software\Microsoft\Office\Teams" "PreventFirstLaunchAfterInstall" 1)
            foreach ($rk in $runKeys) {
                if ($rk.Path -and $rk.Path -match '(?i)CurrentVersion\\Run$' -and (Test-Path -LiteralPath $rk.Path)) {
                    Remove-ItemProperty -LiteralPath $rk.Path -Name "com.squirrel.Teams.Teams" -Force -EA SilentlyContinue
                }
            }
        } catch {}

        Write-Status -Name "Demarrage Windows" -Status OK -Detail ("desactives=$disabled  conserves/on~$kept")
        if ($report.Count -gt 0) {
            $ex = @($report | Select-Object -First 8)
            Write-Host ("  " + $UI.VLine + "  Exemples : " + ($ex -join " | ")) -ForegroundColor DarkGray
            Write-Host ("  " + $UI.VLine + "  Raccourcis retires → $backupRoot") -ForegroundColor DarkGray
        }
        Write-Log ("Startup cleanup: " + ($report -join "; ")) -Level INFO
    } catch {
        Write-Status -Name "Demarrage Windows" -Status FAIL -Detail ("step=$step · $($_.Exception.Message)")
        Write-Log "Startup FATAL at $step : $($_.Exception.Message)" -Level ERR
    }
}

function Set-ProvDesktopOrganize {
    <# Delegue au script standalone desktop-organize.ps1 (aussi dispo via le menu start.bat). #>
    $org = Join-Path $ScriptDir "desktop-organize.ps1"
    if (-not (Test-Path -LiteralPath $org)) {
        Write-Status -Name "Bureau" -Status FAIL -Detail "desktop-organize.ps1 introuvable"
        return
    }
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $org -NoPause
        if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) {
            Write-Status -Name "Bureau" -Status FAIL -Detail "exit $LASTEXITCODE"
        } else {
            Write-Status -Name "Bureau" -Status INFO -Detail "voir resultat ci-dessus" -NoCount
        }
    } catch {
        Write-Status -Name "Bureau" -Status FAIL -Detail $_.Exception.Message
    }
}

# ============================================================================
#  BARRE DES TACHES
#  Explorateur + Chrome + Excel + Word + Discord - tout le reste retire
# ============================================================================

function Find-Shortcut {
    param([string[]]$Names)
    $roots = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs",
        "$env:LOCALAPPDATA\Microsoft\Windows\Start Menu\Programs"
    )
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($n in $Names) {
            $hit = Get-ChildItem -Path $root -Filter $n -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($hit) { return $hit.FullName }
        }
    }
    return $null
}

function New-TaskbarShortcut {
    param([string]$TargetPath, [string]$LinkPath, [string]$Args = "")
    $w = New-Object -ComObject WScript.Shell
    $s = $w.CreateShortcut($LinkPath)
    $s.TargetPath = $TargetPath
    if ($Args) { $s.Arguments = $Args }
    $s.WorkingDirectory = Split-Path $TargetPath -Parent
    $s.Save()
}

function ConvertTo-ProvXmlPath {
    param([string]$PathValue)
    if (-not $PathValue) { return "" }
    return ($PathValue -replace "&", "&amp;" -replace "'", "&apos;")
}

function Set-ProvTaskbar {
    Write-Status -Name "Barre des taches" -Status RUN -Detail ("reset " + $UI.Arrow + " Explorateur / Chrome / Excel / Word / Discord / AnyDesk") -NoCount

    $chromeExe = Find-ChromeExe
    $wordExe = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    $excelExe = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    $anydeskExe = @(
        "${env:ProgramFiles}\AnyDesk\AnyDesk.exe",
        "${env:ProgramFiles(x86)}\AnyDesk\AnyDesk.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    $chromeLnk  = Find-Shortcut -Names @("Google Chrome.lnk")
    $wordLnk    = Find-Shortcut -Names @("Word.lnk", "Microsoft Word.lnk")
    $excelLnk   = Find-Shortcut -Names @("Excel.lnk", "Microsoft Excel.lnk")
    $discordLnk = Find-Shortcut -Names @("Discord.lnk")
    $anydeskLnk = Find-Shortcut -Names @("AnyDesk.lnk")
    $explorerLnk = @(
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\File Explorer.lnk",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\System Tools\File Explorer.lnk",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\File Explorer.lnk"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    $pinDir = Join-Path $env:TEMP ("{0}-taskbar-pins" -f $ProvConfig.FilePrefix)
    New-Item -ItemType Directory -Path $pinDir -Force | Out-Null

    if (-not $chromeLnk -and $chromeExe) {
        $chromeLnk = Join-Path $pinDir "Google Chrome.lnk"
        New-TaskbarShortcut -TargetPath $chromeExe -LinkPath $chromeLnk
    }
    if (-not $wordLnk -and $wordExe) {
        $wordLnk = Join-Path $pinDir "Word.lnk"
        New-TaskbarShortcut -TargetPath $wordExe -LinkPath $wordLnk
    }
    if (-not $excelLnk -and $excelExe) {
        $excelLnk = Join-Path $pinDir "Excel.lnk"
        New-TaskbarShortcut -TargetPath $excelExe -LinkPath $excelLnk
    }
    if (-not $discordLnk) {
        $discordUpdate = Join-Path $env:LOCALAPPDATA "Discord\Update.exe"
        if (Test-Path $discordUpdate) {
            $discordLnk = Join-Path $pinDir "Discord.lnk"
            New-TaskbarShortcut -TargetPath $discordUpdate -LinkPath $discordLnk -Args "--processStart Discord.exe"
        }
    }
    if (-not $anydeskLnk -and $anydeskExe) {
        $anydeskLnk = Join-Path $pinDir "AnyDesk.lnk"
        New-TaskbarShortcut -TargetPath $anydeskExe -LinkPath $anydeskLnk
    }
    if (-not $explorerLnk) {
        $explorerLnk = Join-Path $pinDir "File Explorer.lnk"
        New-TaskbarShortcut -TargetPath "$env:WINDIR\explorer.exe" -LinkPath $explorerLnk
    }

    # LayoutModification.xml - PinListPlacement=Replace = remplace TOUTE la barre
    $shellDir = Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Shell"
    New-Item -ItemType Directory -Path $shellDir -Force | Out-Null

    $pinsXml = New-Object System.Collections.Generic.List[string]
    [void]$pinsXml.Add('        <taskbar:DesktopApp DesktopApplicationID="Microsoft.Windows.Explorer" />')

    foreach ($pair in @(
        @{ Lnk = $chromeLnk;  Label = "Chrome" }
        @{ Lnk = $excelLnk;   Label = "Excel" }
        @{ Lnk = $wordLnk;    Label = "Word" }
        @{ Lnk = $discordLnk; Label = "Discord" }
        @{ Lnk = $anydeskLnk; Label = "AnyDesk" }
    )) {
        if ($pair.Lnk -and (Test-Path $pair.Lnk)) {
            $safe = ConvertTo-ProvXmlPath -PathValue $pair.Lnk
            [void]$pinsXml.Add("        <taskbar:DesktopApp DesktopApplicationLinkPath=`"$safe`" />")
            Write-Log ("Pin : $($pair.Label) " + $UI.Arrow + " $($pair.Lnk)") -Level INFO
        } else {
            Write-Status -Name "Pin $($pair.Label)" -Status FAIL -Detail "raccourci introuvable" -NoCount
            Write-Log "Pin manquant : $($pair.Label)" -Level WARN
        }
    }

    $layout = @"
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
$($pinsXml -join "`n")
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
"@
    $layoutPath = Join-Path $shellDir "LayoutModification.xml"
    Set-Content -Path $layoutPath -Value $layout -Encoding UTF8

    try {
        $polEx = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
        if (-not (Test-Path $polEx)) { New-Item -Path $polEx -Force | Out-Null }
        $machineLayout = Join-Path $env:ProgramData (Join-Path $ProvConfig.WorkFolderName "LayoutModification.xml")
        New-Item -ItemType Directory -Path (Split-Path $machineLayout) -Force | Out-Null
        Copy-Item $layoutPath $machineLayout -Force
        Set-ItemProperty -Path $polEx -Name "StartLayoutFile" -Value $machineLayout -Type String -Force
        Set-ItemProperty -Path $polEx -Name "LockedStartLayout" -Value 0 -Type DWord -Force
    } catch {
        Write-Log "Policy StartLayout : $($_.Exception.Message)" -Level WARN
    }

    # Complement : vider + recopier User Pinned\TaskBar
    $pinnedDir = Join-Path $env:APPDATA "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    New-Item -ItemType Directory -Path $pinnedDir -Force | Out-Null
    Get-ChildItem $pinnedDir -Filter "*.lnk" -EA SilentlyContinue | Remove-Item -Force -EA SilentlyContinue

    $i = 0
    foreach ($lnk in @($explorerLnk, $chromeLnk, $excelLnk, $wordLnk, $discordLnk, $anydeskLnk)) {
        if ($lnk -and (Test-Path $lnk)) {
            $i++
            Copy-Item $lnk (Join-Path $pinnedDir ("{0:D2}_{1}" -f $i, [IO.Path]::GetFileName($lnk))) -Force -EA SilentlyContinue
        }
    }

    $tb = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    if (Test-Path $tb) {
        foreach ($n in @("Favorites","FavoritesResolve","FavoritesVersion","FavoritesChanges")) {
            Remove-ItemProperty -Path $tb -Name $n -EA SilentlyContinue
        }
    }

    try {
        # Redemarre le shell sans laisser une fenetre Explorateur ouverte
        Get-Process explorer -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
        Start-Sleep -Milliseconds 1200
        if (-not (Get-Process explorer -EA SilentlyContinue)) {
            # Relance le shell uniquement (pas une fenetre dossier)
            Start-Process -FilePath "$env:WINDIR\explorer.exe"
            Start-Sleep -Seconds 2
        } else {
            Start-Sleep -Seconds 1
        }
        # Ferme les fenetres Explorateur eventuelles (garde le shell / barre des taches)
        try {
            $shellApp = New-Object -ComObject Shell.Application
            foreach ($w in @($shellApp.Windows())) {
                try {
                    $doc = $w.Document
                    if ($null -eq $doc) { continue }
                    # Fenetre dossier / This PC — pas IE/Edge legacy sans LocationURL dossier
                    if ($w.FullName -match '(?i)explorer\.exe') {
                        $w.Quit()
                    }
                } catch {}
            }
        } catch {
            Write-Log "Close Explorer windows: $($_.Exception.Message)" -Level INFO
        }
        Write-Status -Name "Barre des taches" -Status OK -Detail ("Explorateur " + $UI.DotMid + " Chrome " + $UI.DotMid + " Excel " + $UI.DotMid + " Word " + $UI.DotMid + " Discord " + $UI.DotMid + " AnyDesk")
    } catch {
        Write-Status -Name "Barre des taches" -Status FAIL -Detail $_.Exception.Message
    }
}


# ============================================================================
#  IDENTITE POSTE + COMPTE LOCAL + RESEAU + ANYDESK + BLOAT + MAJ + RAPPORT
# ============================================================================

function Get-ProvAsciiName {
    param([string]$Text)
    return (ConvertTo-ProvEmailLocalPart $Text)
}

function ConvertTo-ProvTitleCase {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = $Text.Trim()
    if ($t.Length -eq 1) { return $t.ToUpperInvariant() }
    # Tirets / espaces : chaque segment
    $parts = $t -split '([\s\-]+)'
    $out = foreach ($seg in $parts) {
        if ($seg -match '^[\s\-]+$') { $seg }
        elseif ($seg.Length -eq 0) { "" }
        elseif ($seg.Length -eq 1) { $seg.ToUpperInvariant() }
        else { $seg.Substring(0,1).ToUpperInvariant() + $seg.Substring(1).ToLowerInvariant() }
    }
    return (-join $out)
}

function Test-ProvLooksLikePersonName {
    <# Evite les faux positifs type username CamelCase (JanedoeXY -> XY). #>
    param([string]$Prenom, [string]$Nom)
    if ([string]::IsNullOrWhiteSpace($Prenom) -or [string]::IsNullOrWhiteSpace($Nom)) { return $false }
    $p = $Prenom.Trim()
    $n = $Nom.Trim()
    # Trop court = bruit (UD, PC, …)
    if ($p.Length -lt 2 -or $n.Length -lt 3) { return $false }
    # Uniquement lettres / tirets / apostrophes / espaces
    if ($p -notmatch "^[\p{L}'][\p{L}'\-]*$" ) { return $false }
    if ($n -notmatch "^[\p{L}'][\p{L}'\- ]*$") { return $false }
    # Prenom tout colle + nom 2 lettres = souvent username casse
    if ($p.Length -ge 10 -and $n.Length -le 2) { return $false }
    return $true
}

function Get-ProvNameFromEmailLocal {
    <#
    Convention : jdoe@<EmailDomain> -> Nom=Doe (1ere lettre = initiale prenom).
    Ne invente PAS le prenom (T seul est inutile).
    #>
    param([string]$Email, [string]$Domain)
    $r = [ordered]@{ Prenom = ""; Nom = "" }
    if ([string]::IsNullOrWhiteSpace($Email) -or $Email -notmatch '@') { return [pscustomobject]$r }
    $local = (($Email -split '@')[0]).ToLowerInvariant()
    $dom = (($Email -split '@')[1]).ToLowerInvariant()
    if ($Domain -and $dom -ne $Domain.ToLowerInvariant()) {
        # Hors domaine de l'organisation : on ne parse pas la convention
        return [pscustomobject]$r
    }
    if ($local.Length -lt 3 -or $local -notmatch '^[a-z]') { return [pscustomobject]$r }
    # local = initiale + nom (ex. jdoe)
    $r.Nom = ConvertTo-ProvTitleCase ($local.Substring(1))
    return [pscustomobject]$r
}

function Get-ProvSuggestedIdentity {
    <#
    Pre-remplissage fiable :
      1) Email : UPN Entra / Office
      2) Display name propre : Office FriendlyName / Identities / FullName
      3) Nom via convention email (jdoe -> Doe) — sans inventer le prenom
      4) Username CamelCase SEULEMENT si qualite OK (rejette JanedoeXY)
    #>
    $sug = [ordered]@{ Prenom = ""; Nom = ""; Email = ""; Source = "" }
    $domain = [string]$ProvConfig.EmailDomain

    # --- Email ---
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
            $id = Get-ItemProperty -LiteralPath $ip -ErrorAction SilentlyContinue
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
            foreach ($sub in @(Get-ChildItem -LiteralPath $idDir -ErrorAction SilentlyContinue)) {
                $props = Get-ItemProperty -LiteralPath $sub.PSPath -ErrorAction SilentlyContinue
                if (-not $props) { continue }
                foreach ($prop in @("EmailAddress", "EmailAddresses", "UserName")) {
                    if (-not ($props.PSObject.Properties.Name -contains $prop)) { continue }
                    $cand = [string]$props.$prop
                    if ($cand -match '([^\s;]+@[^\s;]+)') {
                        $sug.Email = $Matches[1].ToLowerInvariant()
                        if ($sug.Source -ne "UPN") { $sug.Source = "Office" }
                    }
                }
                foreach ($prop in @("FriendlyName", "DisplayName", "Name")) {
                    if (-not ($props.PSObject.Properties.Name -contains $prop)) { continue }
                    $v = [string]$props.$prop
                    if ($v -match '\S+\s+\S+') { $officeDisplay = $v; break }
                }
            }
        }
    } catch {}

    # --- Display name (Office puis profil Windows) ---
    $nameApplied = $false
    $displayCandidates = New-Object System.Collections.Generic.List[string]
    if ($officeDisplay) { [void]$displayCandidates.Add($officeDisplay) }
    try {
        $lu = Get-LocalUser -Name $env:USERNAME -ErrorAction SilentlyContinue
        if ($lu -and $lu.FullName) { [void]$displayCandidates.Add([string]$lu.FullName) }
    } catch {}
    try {
        $ads = [ADSI]"WinNT://$env:COMPUTERNAME/$($env:USERNAME),user"
        if ($ads -and $ads.FullName -and $ads.FullName.Value) {
            [void]$displayCandidates.Add([string]$ads.FullName.Value)
        }
    } catch {}

    foreach ($fn in $displayCandidates) {
        if (-not $fn -or $fn -notmatch '\S+\s+\S+') { continue }
        $parts = $fn.Trim() -split '\s+', 2
        if (-not (Test-ProvLooksLikePersonName -Prenom $parts[0] -Nom $parts[1])) { continue }
        $sug.Prenom = ConvertTo-ProvTitleCase $parts[0]
        $sug.Nom = ConvertTo-ProvTitleCase $parts[1]
        if ($fn -eq $officeDisplay) { $sug.Source = "Office" }
        elseif (-not $sug.Source -or $sug.Source -eq "UPN") { $sug.Source = "Profil" }
        $nameApplied = $true
        break
    }

    # --- Convention email : jdoe -> Nom Doe (jamais inventer le prenom) ---
    if ($sug.Email) {
        $fromMail = Get-ProvNameFromEmailLocal -Email $sug.Email -Domain $domain
        if ($fromMail.Nom) {
            $mailNomAscii = ConvertTo-ProvEmailLocalPart $fromMail.Nom
            $curNomAscii = ""
            if ($sug.Nom) { $curNomAscii = ConvertTo-ProvEmailLocalPart $sug.Nom }

            $nameOk = $false
            if ($sug.Prenom -and $sug.Nom) {
                $nameOk = Test-ProvLooksLikePersonName -Prenom $sug.Prenom -Nom $sug.Nom
            }
            $coherent = $false
            if ($curNomAscii -and $mailNomAscii) {
                if ($mailNomAscii -eq $curNomAscii) { $coherent = $true }
                elseif ($mailNomAscii.EndsWith($curNomAscii)) { $coherent = $true }
                elseif ($curNomAscii.EndsWith($mailNomAscii)) { $coherent = $true }
            }

            # Prenom type "Janedo" qui englobe le local email -> jeter
            if ($sug.Prenom) {
                $pA = ConvertTo-ProvEmailLocalPart $sug.Prenom
                $local = (($sug.Email -split '@')[0]).ToLowerInvariant()
                if ($pA.Length -ge 6 -and $local.Contains($pA)) {
                    $sug.Prenom = ""
                    $nameOk = $false
                }
            }

            if (-not $nameOk -or -not $coherent) {
                $sug.Nom = $fromMail.Nom
                if ($sug.Prenom -and -not (Test-ProvLooksLikePersonName -Prenom $sug.Prenom -Nom $sug.Nom)) {
                    $sug.Prenom = ""
                }
                if (-not $nameApplied -or -not $coherent) {
                    if ($sug.Source -and $sug.Source -ne "Email") {
                        $sug.Source = "$($sug.Source)+Email"
                    } else {
                        $sug.Source = "Email"
                    }
                }
            }
        }
    }

    # --- Username CamelCase strict (dernier recours) ---
    if ((-not $sug.Prenom -or -not $sug.Nom) -and $env:USERNAME) {
        $u = $env:USERNAME
        $candP = $null
        $candN = $null
        # Exige majuscule + au moins 3 lettres chaque cote (rejette …UD)
        if ($u -match '^([A-Z][a-z]{2,})([A-Z][A-Za-z\-]{2,})$') {
            $candP = $Matches[1]
            $candN = $Matches[2]
        } elseif ($u -match '^([A-Za-z]{2,})\.([A-Za-z]{2,})$') {
            $candP = $Matches[1]
            $candN = $Matches[2]
        }
        if ($candP -and $candN -and (Test-ProvLooksLikePersonName -Prenom $candP -Nom $candN)) {
            if (-not $sug.Prenom) { $sug.Prenom = ConvertTo-ProvTitleCase $candP }
            if (-not $sug.Nom) { $sug.Nom = ConvertTo-ProvTitleCase $candN }
            if (-not $sug.Source) { $sug.Source = "Username" }
        }
    }

    # Filet de securite
    if ($sug.Prenom -and $sug.Nom) {
        if (-not (Test-ProvLooksLikePersonName -Prenom $sug.Prenom -Nom $sug.Nom)) {
            $sug.Prenom = ""
            if ($sug.Email) {
                $fromMail = Get-ProvNameFromEmailLocal -Email $sug.Email -Domain $domain
                if ($fromMail.Nom) { $sug.Nom = $fromMail.Nom; $sug.Source = "Email" }
            } else {
                $sug.Nom = ""
            }
        }
    }

    return [pscustomobject]$sug
}

function Get-ProvReportRecipients {
    <# Normalise ReportTo (string / array) en liste d'emails. #>
    $raw = $ProvConfig.ReportTo
    $list = @()
    if ($null -eq $raw) { return @() }
    if ($raw -is [System.Array]) {
        foreach ($x in $raw) {
            if ($x) { $list += ($x.ToString().Trim()) }
        }
    } else {
        foreach ($x in ($raw.ToString() -split '[;,]')) {
            $t = $x.Trim()
            if ($t) { $list += $t }
        }
    }
    return @($list | Where-Object { $_ -match '^[^@\s]+@[^@\s]+$' } | Select-Object -Unique)
}

function Set-ProvComputerIdentity {
    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " Identite du poste " + ($UI.HLine * 34)) -ForegroundColor Magenta
    Write-Host ("  " + $UI.VLine + "  Nom actuel : $env:COMPUTERNAME") -ForegroundColor Gray
    $nameEx = @(Get-ProvHostnameExamples -Config $ProvConfig)
    Write-Host ("  " + $UI.VLine + "  Format : $($ProvConfig.OrgPrefix)-<SITE>-<ITEM>-<NUM>  (max 15 car.)") -ForegroundColor DarkGray
    $siteParts = @(@($ProvConfig.Sites) | ForEach-Object { "$($_.Code) $($_.Name)" })
    for ($si = 0; $si -lt $siteParts.Count; $si += 3) {
        $siteLbl = if ($si -eq 0) { "  SITE : " } else { "         " }
        $siteChunk = $siteParts[$si..([Math]::Min($si + 2, $siteParts.Count - 1))] -join " | "
        Write-Host ("  " + $UI.VLine + $siteLbl + $siteChunk) -ForegroundColor DarkGray
    }
    $itemParts = @(@($ProvConfig.ItemTypes) | ForEach-Object { if ($_.Hint) { "$($_.Code) ($($_.Hint))" } else { [string]$_.Code } })
    Write-Host ("  " + $UI.VLine + "  ITEM : " + ($itemParts -join " | ")) -ForegroundColor DarkGray
    Write-Host ("  " + $UI.VLine + "  Ex.  : " + ($nameEx -join "  /  ")) -ForegroundColor Cyan
    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
    Write-Host ""

    # Reprise session precedente (si script coupe a mi-chemin)
    $saved = Read-ProvDeployState
    $useSaved = $false
    if ($saved -and ($saved.ComputerName -or $saved.Prenom -or $saved.Nom -or $saved.Email)) {
        $when = if ($saved.UpdatedAt) { $saved.UpdatedAt } else { "?" }
        Write-Host ("  " + $UI.VLine + "  Session precedente trouvee ($when)") -ForegroundColor Yellow
        if ($saved.ComputerName) { Write-Host ("  " + $UI.VLine + "    PC    : $($saved.ComputerName)") -ForegroundColor DarkYellow }
        if ($saved.Prenom -or $saved.Nom) {
            Write-Host ("  " + $UI.VLine + "    User  : $($saved.Prenom) $($saved.Nom)") -ForegroundColor DarkYellow
        }
        if ($saved.Email) { Write-Host ("  " + $UI.VLine + "    Email : $($saved.Email)") -ForegroundColor DarkYellow }
        Write-Host ""
        if (Read-ProvYesNo -Prompt "Reprendre ces infos ?" -Default Yes) {
            $useSaved = $true
            Write-Status -Name "Session" -Status OK -Detail "reprise des infos sauvees" -NoCount
        } else {
            Clear-ProvDeployState
            $saved = $null
            Write-Status -Name "Session" -Status INFO -Detail "infos precedentes ignorees" -NoCount
        }
        Write-Host ""
    }

    $maxRenameTries = 3
    $renamed = $false
    $nameRegex = Get-ProvHostnameRegex -Config $ProvConfig
    if ($ProvConfig.IsExample) {
        # config.example.psd1 = valeurs fictives : on ne renomme pas le poste
        Write-Status -Name "Rename PC" -Status SKIP -Detail "config.example.psd1 — creer config.local.psd1 pour activer le renommage"
        $maxRenameTries = 0
        $script:ProvComputerName = $env:COMPUTERNAME
        $renamed = $true
    }
    for ($try = 1; $try -le $maxRenameTries; $try++) {
        $defaultName = $env:COMPUTERNAME
        if ($useSaved -and $saved.ComputerName) { $defaultName = [string]$saved.ComputerName }
        $prompt = if ($try -eq 1) {
            "     Nouveau nom du PC [Entree = $defaultName]"
        } else {
            "     Nouveau nom (essai $try/$maxRenameTries) [Entree = $defaultName]"
        }
        $newName = (Read-Host $prompt).Trim()
        if ([string]::IsNullOrWhiteSpace($newName)) {
            $newName = $defaultName
        }
        $newName = ($newName -replace '[^A-Za-z0-9\-]', '').ToUpperInvariant()

        # Entree sur nom actuel = garder sans rename
        if ($newName -eq $env:COMPUTERNAME -or [string]::IsNullOrWhiteSpace($newName)) {
            Write-Status -Name "Rename PC" -Status SKIP -Detail $env:COMPUTERNAME
            $script:ProvComputerName = $env:COMPUTERNAME
            Save-ProvDeployState -Patch @{ ComputerName = $script:ProvComputerName }
            $renamed = $true
            break
        }
        if ($newName.Length -lt 1 -or $newName.Length -gt 15) {
            Write-Status -Name "Rename PC" -Status FAIL -Detail "nom invalide (1-15 car.) — Windows limite a 15" -NoCount
            continue
        }
        if ($newName -notmatch $nameRegex) {
            Write-Status -Name "Rename PC" -Status FAIL -Detail ("format {0}-SITE-ITEM-NUM (ex. {1})" -f $ProvConfig.OrgPrefix, $nameEx[0]) -NoCount
            continue
        }
        # Deja renomme lors d'un run precedent (reboot pas encore fait) : nom cible != actuel
        if ($newName -eq $env:COMPUTERNAME) {
            Write-Status -Name "Rename PC" -Status SKIP -Detail "inchange"
            $script:ProvComputerName = $env:COMPUTERNAME
            Save-ProvDeployState -Patch @{ ComputerName = $script:ProvComputerName }
            $renamed = $true
            break
        }
        try {
            Rename-Computer -NewName $newName -Force -ErrorAction Stop
            $script:ProvComputerName = $newName
            Write-Status -Name "Rename PC" -Status OK -Detail "$newName (effectif apres redemarrage)"
            Save-ProvDeployState -Patch @{ ComputerName = $script:ProvComputerName }
            $renamed = $true
            break
        } catch {
            # Souvent : deja programme au reboot precedent
            if ($_.Exception.Message -match 'already|deja|pending|redemarr') {
                $script:ProvComputerName = $newName
                Write-Status -Name "Rename PC" -Status OK -Detail "$newName (deja en attente de reboot)"
                Save-ProvDeployState -Patch @{ ComputerName = $script:ProvComputerName }
                $renamed = $true
                break
            }
            Write-Status -Name "Rename PC" -Status FAIL -Detail $_.Exception.Message -NoCount
        }
    }
    if (-not $renamed) {
        Write-Status -Name "Rename PC" -Status FAIL -Detail "abandon apres $maxRenameTries essais — nom conserve: $env:COMPUTERNAME"
        $script:ProvComputerName = $env:COMPUTERNAME
        if ($useSaved -and $saved.ComputerName) { $script:ProvComputerName = [string]$saved.ComputerName }
        Save-ProvDeployState -Patch @{ ComputerName = $script:ProvComputerName }
    }

    Write-Host ""
    $sug = Get-ProvSuggestedIdentity
    # Priorite : session sauvee > suggestion Office/UPN
    if ($useSaved) {
        if ($saved.Prenom) { $sug.Prenom = [string]$saved.Prenom; $sug.Source = "Session" }
        if ($saved.Nom) { $sug.Nom = [string]$saved.Nom; if ($sug.Source -ne "Session") { $sug.Source = "Session" } }
        if ($saved.Email) { $sug.Email = [string]$saved.Email }
    }
    if ($sug.Prenom -or $sug.Nom -or $sug.Email) {
        $who = (@($sug.Prenom, $sug.Nom) | Where-Object { $_ }) -join " "
        if (-not $who) { $who = "(nom depuis email)" }
        Write-Host ("  " + $UI.VLine + "  Suggestion ($($sug.Source)) : $who  $($sug.Email)") -ForegroundColor DarkCyan
        Write-Host ("  " + $UI.VLine + "  Entree = accepter la suggestion") -ForegroundColor DarkGray
    }

    $pPrompt = "     Prenom utilisateur"
    if ($sug.Prenom) { $pPrompt += " [$($sug.Prenom)]" }
    $nPrompt = "     Nom utilisateur"
    if ($sug.Nom) { $nPrompt += " [$($sug.Nom)]" }
    $pIn = (Read-Host $pPrompt).Trim()
    $nIn = (Read-Host $nPrompt).Trim()
    if ([string]::IsNullOrWhiteSpace($pIn) -and $sug.Prenom) { $pIn = $sug.Prenom }
    if ([string]::IsNullOrWhiteSpace($nIn) -and $sug.Nom) { $nIn = $sug.Nom }
    $script:ProvPrenom = $pIn
    $script:ProvNom = $nIn

    $p = Get-ProvAsciiName $script:ProvPrenom
    $n = Get-ProvAsciiName $script:ProvNom
    if ($p.Length -ge 1 -and $n.Length -ge 1) {
        $generated = ($p.Substring(0,1) + $n) + "@$($ProvConfig.EmailDomain)"
        $script:ProvEmail = $generated
        if ($useSaved -and $saved.Email) {
            $script:ProvEmail = [string]$saved.Email
        } elseif ($sug.Email -and $sug.Email -match [regex]::Escape("@$($ProvConfig.EmailDomain)$")) {
            if ($sug.Email.ToLowerInvariant() -ne $generated.ToLowerInvariant()) {
                Write-Status -Name "Identite" -Status INFO -Detail "Office/UPN=$($sug.Email) — email genere=$generated" -NoCount
            }
        }
        # Confirmer / corriger email (prefill session ou genere)
        $ePrompt = "     Email [$($script:ProvEmail)]"
        $eIn = (Read-Host $ePrompt).Trim()
        if (-not [string]::IsNullOrWhiteSpace($eIn)) { $script:ProvEmail = $eIn }

        Write-Status -Name "Identite" -Status OK -Detail "$($script:ProvPrenom) $($script:ProvNom) / $($script:ProvEmail)"
        Save-ProvDeployState
    } else {
        Write-Status -Name "Identite" -Status FAIL -Detail "prenom/nom incomplets — email non genere"
        Save-ProvDeployState
    }
}

function Test-ProvAzureAdJoined {
    try {
        $status = & dsregcmd.exe /status 2>$null | Out-String
        return [bool]($status -match "AzureAdJoined\s*:\s*YES")
    } catch {
        return $false
    }
}

function New-ProvComplexPassword {
    <# Garantit majuscule + minuscule + chiffre + special (politique Windows). #>
    $upper = [char](Get-Random -Minimum 65 -Maximum 91)
    $lower = [char](Get-Random -Minimum 97 -Maximum 123)
    $digit = [char](Get-Random -Minimum 48 -Maximum 58)
    $special = @('!','@','#','$','%','*','+','=','?') | Get-Random
    $rest = -join ((48..57 + 65..90 + 97..122) | Get-Random -Count 8 | ForEach-Object { [char]$_ })
    $chars = @($upper, $lower, $digit, $special) + $rest.ToCharArray()
    return -join ($chars | Get-Random -Count $chars.Count)
}

function New-ProvLocalUser {
    <#
    Compte local Windows — desactive par defaut si le poste est deja Entra ID / Intune
    (l'utilisateur se connecte avec jdoe@<EmailDomain>, pas besoin d'un compte local).
    #>
    $aad = Test-ProvAzureAdJoined
    $who = $env:USERNAME
    Write-Host ""
    if ($aad) {
        Write-Host ("  " + $UI.VLine + "  Poste Entra ID / Intune detecte — compte cloud : $who") -ForegroundColor Gray
        Write-Host ("  " + $UI.VLine + "  Un compte local n'est en general PAS necessaire.") -ForegroundColor DarkGray
        $ans = (Read-Host "     Creer quand meme un compte local Windows ? [o/N]").Trim().ToLowerInvariant()
        if ($ans -notin @('o','oui','y','yes')) {
            Write-Status -Name "Compte local" -Status SKIP -Detail "Intune/Entra — non demande"
            return
        }
    } else {
        $ans = (Read-Host "     Creer un compte local Windows (hors Intune) ? [o/N]").Trim().ToLowerInvariant()
        if ($ans -notin @('o','oui','y','yes')) {
            Write-Status -Name "Compte local" -Status SKIP -Detail "non demande"
            return
        }
    }

    if ([string]::IsNullOrWhiteSpace($script:ProvPrenom)) {
        Write-Status -Name "Compte local" -Status SKIP -Detail "prenom absent"
        return
    }
    $userName = Get-ProvAsciiName $script:ProvPrenom
    if ($userName.Length -lt 1) { $userName = "utilisateur" }
    if ($userName.Length -gt 20) { $userName = $userName.Substring(0, 20) }

    Write-Status -Name "Compte local" -Status RUN -Detail "login=$userName" -NoCount

    $existing = Get-LocalUser -Name $userName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Status -Name "Compte local" -Status SKIP -Detail "existe deja : $userName"
        return
    }

    Write-Host ("     Mot de passe pour le compte local '$userName'") -ForegroundColor Gray
    Write-Host ("     (laisser vide = generer un mot de passe temporaire complexe)") -ForegroundColor DarkGray
    $sec1 = Read-Host "     Mot de passe" -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec1)
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) | Out-Null

    if ([string]::IsNullOrWhiteSpace($plain)) {
        $plain = New-ProvComplexPassword
        Write-Host ("     MDP temporaire : $plain") -ForegroundColor Yellow
        Write-Host ("     (a changer a la 1ere connexion — notez-le)") -ForegroundColor DarkYellow
        Write-Log "Compte local $userName MDP temporaire genere (voir console)" -Level WARN
    }
    $sec = ConvertTo-SecureString $plain -AsPlainText -Force

    try {
        $full = ("{0} {1}" -f $script:ProvPrenom, $script:ProvNom).Trim()
        New-LocalUser -Name $userName -Password $sec -FullName $full -Description ("Utilisateur " + $ProvConfig.OrgName) -PasswordNeverExpires:$false -ErrorAction Stop | Out-Null
        Add-LocalGroupMember -Group "Users" -Member $userName -ErrorAction SilentlyContinue
        Write-Status -Name "Compte local" -Status OK -Detail $userName
    } catch {
        Write-Status -Name "Compte local" -Status FAIL -Detail $_.Exception.Message
    }
}

function Test-ProvOfflineRootValid {
    <# Un dossier "offline" contient au moins un marqueur connu (config OfflineMarkers). #>
    param([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path -PathType Container)) { return $false }
    $markers = @($ProvConfig.OfflineMarkers)
    foreach ($m in $markers) {
        if (Test-Path -LiteralPath (Join-Path $Path $m)) { return $true }
    }
    # zip offline du client camera (config CameraClientOfflineZip)
    if (-not $ProvConfig.CameraClientOfflineZip) { return $false }
    $z = Get-ChildItem -LiteralPath $Path -Filter $ProvConfig.CameraClientOfflineZip -File -EA SilentlyContinue | Select-Object -First 1
    return [bool]$z
}

function Get-ProvOfflineRoots {
    <#
    Ordre de recherche des assets offline :
      1) Cle USB / volumes amovibles  (scripts\offline, offline, usb_pro\downloads)
      2) Autres lecteurs (USB parfois monte en Fixed) sauf systeme
      3) A cote du .ps1 (ex. Bureau\scripts\offline)
      4) Bureau / Downloads
    Ainsi : offline reste sur la cle ; le .ps1 peut etre copie sur le Bureau.
    #>
    $ordered = New-Object System.Collections.Generic.List[string]

    function Add-ProvRoot([string]$Candidate, [switch]$RequireMarker) {
        if ([string]::IsNullOrWhiteSpace($Candidate)) { return }
        try {
            if (-not (Test-Path -LiteralPath $Candidate -PathType Container)) { return }
            $full = (Resolve-Path -LiteralPath $Candidate).Path
            if ($ordered.Contains($full)) { return }
            if ($RequireMarker -and -not (Test-ProvOfflineRootValid -Path $full)) { return }
            [void]$ordered.Add($full)
        } catch {}
    }

    $removable = @()
    $other = @()
    try {
        foreach ($d in @(Get-CimInstance -ClassName Win32_LogicalDisk -ErrorAction SilentlyContinue)) {
            if (-not $d.DeviceID) { continue }
            if ($d.DriveType -eq 2) {
                $removable += $d.DeviceID
            } elseif ($d.DriveType -eq 3 -and $d.DeviceID -ne $env:SystemDrive) {
                $other += $d.DeviceID
            }
        }
    } catch {}

    foreach ($letter in ($removable + $other)) {
        # RequireMarker : evite un dossier "offline" aleatoire sur un disque data
        Add-ProvRoot (Join-Path $letter "scripts\offline") -RequireMarker
        Add-ProvRoot (Join-Path $letter "offline") -RequireMarker
        Add-ProvRoot (Join-Path $letter "usb_pro\downloads") -RequireMarker
        Add-ProvRoot (Join-Path $letter "scripts") -RequireMarker
    }

    # A cote du script : seulement si offline valide (ne cree JAMAIS le dossier)
    Add-ProvRoot (Join-Path $ScriptDir "offline") -RequireMarker
    Add-ProvRoot $ScriptDir
    Add-ProvRoot (Join-Path $ScriptDir "..\usb_pro\downloads") -RequireMarker
    Add-ProvRoot (Join-Path $ScriptDir "usb_pro\downloads") -RequireMarker
    Add-ProvRoot (Join-Path $ScriptDir "..\downloads") -RequireMarker
    Add-ProvRoot (Join-Path $ScriptDir "..\offline") -RequireMarker

    $desk = [Environment]::GetFolderPath("Desktop")
    $docs = [Environment]::GetFolderPath("MyDocuments")
    $dl   = Join-Path $env:USERPROFILE "Downloads"
    foreach ($base in @($desk, $docs, $dl)) {
        if (-not $base) { continue }
        Add-ProvRoot (Join-Path $base "scripts\offline") -RequireMarker
        Add-ProvRoot (Join-Path $base "offline") -RequireMarker
        Add-ProvRoot (Join-Path $base "usb_pro\downloads") -RequireMarker
    }

    return @($ordered)
}

function Test-ProvOfflinePreflight {
    <#
    Controles AVANT winget/Office : la cle a-t-elle le minimum metier ?
    WARN si partiel ; ne bloque pas (operateur decide).
    #>
    Write-Status -Name "Preflight offline" -Status RUN -Detail "controle assets USB..." -NoCount
    $roots = @(Get-ProvOfflineRoots)
    if ($roots.Count -eq 0) {
        Write-Status -Name "Preflight offline" -Status FAIL -Detail "aucune source offline (brancher la cle)"
        return $false
    }
    Write-Status -Name "Preflight source" -Status INFO -Detail ($roots[0]) -NoCount

    $checks = @()
    $pa = $ProvConfig.PortableApp
    if ($pa -and $pa.Folder) {
        $checks += @{ Name = [string]$pa.Folder; Rel = @($pa.Folder); Dir = $true; Critical = $true }
    }
    $lob = $ProvConfig.LobApp
    if ($lob -and $lob.SourceFolders) {
        $lobRel = if ($lob.PreflightFolders) { @($lob.PreflightFolders) } else { @($lob.SourceFolders) }
        $checks += @{ Name = [string]$lob.Name; Rel = $lobRel; Dir = $true; Critical = $true }
    }
    $checks += @{ Name = "ODT Office";          Rel = @("officedeploymenttool.exe", "odt.exe"); Dir = $false; Critical = $true }
    $checks += @{ Name = "AnyDesk.exe";         Rel = @("AnyDesk.exe"); Dir = $false; Critical = $false }
    $co = $ProvConfig.ClickOnceApp
    if ($co -and $co.AppRefNames) {
        $coLabel = if ($co.PreflightLabel) { [string]$co.PreflightLabel } else { "$($co.Name) ClickOnce" }
        $checks += @{ Name = $coLabel; Rel = @(@($co.AppRefNames)[0]); Dir = $false; Critical = $false }
    }
    $checks += @{ Name = "Drivers imprimante";  Rel = @("printers"); Dir = $true; Critical = $false }
    if ($ProvConfig.CameraClientEnabled) {
        $checks += @{ Name = ("{0} setup" -f (Get-ProvCameraClientName)); Rel = @(); Dir = $false; Critical = $false; CameraClient = $true }
    }

    $criticalMiss = 0
    foreach ($c in $checks) {
        $hit = $null
        if ($c.CameraClient) {
            $ccInner = [string]$ProvConfig.CameraClientExePattern
            if (-not $ccInner) { $ccInner = '.' }
            foreach ($root in $roots) {
                if ($ProvConfig.CameraClientPreflightPattern) {
                    $hit = Get-ChildItem -LiteralPath $root -File -EA SilentlyContinue |
                        Where-Object { $_.Name -match $ProvConfig.CameraClientPreflightPattern } |
                        Select-Object -First 1
                    if ($hit) { break }
                }
                if (-not $ProvConfig.CameraClientOfflineSubdir) { continue }
                $sub = Join-Path $root $ProvConfig.CameraClientOfflineSubdir
                if (Test-Path $sub) {
                    $hit = Get-ChildItem -LiteralPath $sub -Recurse -Include *.exe,*.msi,*.zip -EA SilentlyContinue |
                        Where-Object { $_.Name -match $ccInner -and $_.Name -notmatch '(?i)unins|\.lnk$' } |
                        Select-Object -First 1
                    if ($hit) { break }
                }
            }
            if ($hit) {
                Write-Status -Name $c.Name -Status OK -Detail $hit.Name -NoCount
            } else {
                Write-Status -Name $c.Name -Status INFO -Detail "absent offline — download / winget en fallback" -NoCount
            }
            continue
        }
        if ($c.Dir) {
            $hit = Find-ProvOfflineItem -RelativeNames $c.Rel -Directory
        } else {
            $hit = Find-ProvOfflineFile -Names $c.Rel
        }
        if ($hit) {
            Write-Status -Name $c.Name -Status OK -Detail $hit -NoCount
        } else {
            if ($c.Critical) {
                Write-Status -Name $c.Name -Status FAIL -Detail "MANQUANT (critique)" -NoCount
                $criticalMiss++
            } else {
                Write-Status -Name $c.Name -Status INFO -Detail "absent (optionnel)" -NoCount
            }
        }
    }

    if ($criticalMiss -gt 0) {
        Write-Status -Name "Preflight offline" -Status FAIL -Detail "$criticalMiss asset(s) critique(s) manquant(s)"
        Write-Host ("  " + $UI.Arrow + " Continuer quand meme ? Le metier peut echouer.") -ForegroundColor Yellow
        $ans = (Read-Host "     Continuer malgre preflight FAIL ? [o/N]").Trim().ToLowerInvariant()
        return ($ans -in @('o', 'oui', 'y', 'yes'))
    }
    Write-Status -Name "Preflight offline" -Status OK -Detail "assets critiques presents"
    return $true
}

function Find-ProvOfflineFile {
    <# Cherche un fichier par nom dans les racines offline (USB d'abord). #>
    param([string[]]$Names)
    foreach ($root in @(Get-ProvOfflineRoots)) {
        foreach ($n in $Names) {
            if ([string]::IsNullOrWhiteSpace($n)) { continue }
            $p = Join-Path $root $n
            if (Test-Path -LiteralPath $p -PathType Leaf) { return $p }
        }
    }
    return $null
}

function Find-ProvOfflineItem {
    param(
        [string[]]$RelativeNames,
        [switch]$Directory
    )
    foreach ($root in @(Get-ProvOfflineRoots)) {
        foreach ($rel in $RelativeNames) {
            $p = Join-Path $root $rel
            if ($Directory) {
                if (Test-Path -LiteralPath $p -PathType Container) { return $p }
            } else {
                if (Test-Path -LiteralPath $p -PathType Leaf) { return $p }
            }
        }
    }
    return $null
}

function Copy-ProvDirectoryRobust {
    param([string]$Source, [string]$Destination)
    if (-not (Test-Path -LiteralPath $Source)) { throw "Source introuvable: $Source" }
    $parent = Split-Path -Parent $Destination
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    # robocopy : codes 0-7 = succes partiel/ok
    $null = & robocopy.exe $Source $Destination /E /R:2 /W:2 /NFL /NDL /NJH /NJS /NP
    if ($LASTEXITCODE -ge 8) {
        Copy-Item -LiteralPath $Source -Destination $Destination -Recurse -Force -ErrorAction Stop
    }
    return (Test-Path -LiteralPath $Destination)
}

function New-ProvStartMenuShortcut {
    param(
        [string]$Name,
        [string]$TargetPath,
        [string]$Arguments = "",
        [string]$WorkingDirectory = "",
        [string]$IconLocation = ""
    )
    $dir = Join-Path $env:ProgramData (Join-Path "Microsoft\Windows\Start Menu\Programs" $ProvConfig.StartMenuFolder)
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $lnk = Join-Path $dir ($Name + ".lnk")
    $w = New-Object -ComObject WScript.Shell
    $s = $w.CreateShortcut($lnk)
    $s.TargetPath = $TargetPath
    if ($Arguments) { $s.Arguments = $Arguments }
    if ($WorkingDirectory) { $s.WorkingDirectory = $WorkingDirectory }
    elseif ($TargetPath -and (Test-Path $TargetPath)) { $s.WorkingDirectory = Split-Path $TargetPath -Parent }
    if ($IconLocation) { $s.IconLocation = $IconLocation }
    $s.Save()
    return $lnk
}

function Find-ProvPrinterDriverInfs {
    <#
    INF du pack driver dans offline\printers\ (ordre : config PrinterDriverFolders).
    Ignore les INF exclus par PrinterInfExcludeDirPattern / PrinterInfExcludeNamePattern.
    #>
    $infs = New-Object System.Collections.Generic.List[string]
    $preferDirs = @($ProvConfig.PrinterDriverFolders | Where-Object { $_ })
    if ($preferDirs.Count -eq 0) { $preferDirs = @("printers") }
    $exDir = [string]$ProvConfig.PrinterInfExcludeDirPattern
    $exName = [string]$ProvConfig.PrinterInfExcludeNamePattern
    foreach ($root in @(Get-ProvOfflineRoots)) {
        foreach ($rel in $preferDirs) {
            $dir = Join-Path $root $rel
            if (-not (Test-Path -LiteralPath $dir -PathType Container)) { continue }
            $found = @(Get-ChildItem -LiteralPath $dir -Filter "*.inf" -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object {
                    (-not $exDir -or $_.FullName -notmatch $exDir) -and
                    (-not $exName -or $_.Name -notmatch $exName)
                })
            foreach ($f in $found) {
                if (-not $infs.Contains($f.FullName)) { [void]$infs.Add($f.FullName) }
            }
        }
    }
    return @($infs)
}

function Install-ProvOfflinePrinterDrivers {
    <#
    Stage + install silencieux du pack driver (config PrinterDriverNames) depuis offline\printers via pnputil.
    Necessaire : les drivers constructeur ne sont souvent PAS inbox sur Win11 neuf.
    #>
    $wanted = @($ProvConfig.PrinterDriverNames | Where-Object { $_ })
    if ($wanted.Count -eq 0) {
        Write-Status -Name "Drivers imprimante" -Status SKIP -Detail "aucun PrinterDriverNames dans la config" -NoCount
        return $false
    }
    $drvLabel = [string]$ProvConfig.PrinterDriverLabel
    if (-not $drvLabel) { $drvLabel = $wanted[0] }
    Write-Status -Name "Drivers imprimante" -Status RUN -Detail "$drvLabel (offline)..." -NoCount

    # Deja presents ?
    $already = @()
    foreach ($n in $wanted) {
        if (Get-PrinterDriver -Name $n -ErrorAction SilentlyContinue) { $already += $n }
    }
    if ($already -contains $wanted[0]) {
        Write-Status -Name "Drivers imprimante" -Status SKIP -Detail ($already -join ", ")
        return $true
    }

    $infs = @(Find-ProvPrinterDriverInfs)
    if ($infs.Count -eq 0) {
        Write-Status -Name "Drivers imprimante" -Status INFO -Detail "pas de pack offline\printers — fallback inbox/WU" -NoCount
        return $false
    }

    $staged = 0
    foreach ($inf in $infs) {
        $leaf = Split-Path -Leaf $inf
        try {
            $p = Start-Process -FilePath "$env:SystemRoot\System32\pnputil.exe" `
                -ArgumentList @("/add-driver", $inf, "/install") `
                -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
            # 0 = OK, 3010 = reboot required, 259 = no more items / deja la (varie)
            if ($p.ExitCode -in @(0, 3010)) {
                Write-Status -Name "pnputil" -Status OK -Detail "$leaf (exit $($p.ExitCode))" -NoCount
                $staged++
            } else {
                Write-Status -Name "pnputil" -Status INFO -Detail "$leaf exit $($p.ExitCode)" -NoCount
                Write-Log "pnputil /add-driver $inf exit=$($p.ExitCode)" -Level WARN
                # Retry sans /install (stage store seulement) puis Add-PrinterDriver
                $p2 = Start-Process -FilePath "$env:SystemRoot\System32\pnputil.exe" `
                    -ArgumentList @("/add-driver", $inf) `
                    -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
                if ($p2 -and $p2.ExitCode -in @(0, 3010)) {
                    Write-Status -Name "pnputil" -Status OK -Detail "$leaf staged (exit $($p2.ExitCode))" -NoCount
                    $staged++
                }
            }
        } catch {
            Write-Status -Name "pnputil" -Status FAIL -Detail "$leaf — $($_.Exception.Message)" -NoCount
        }
    }

    foreach ($n in $wanted) {
        if (Get-PrinterDriver -Name $n -ErrorAction SilentlyContinue) { continue }
        try {
            Add-PrinterDriver -Name $n -ErrorAction Stop
            Write-Status -Name "Add-PrinterDriver" -Status OK -Detail $n -NoCount
        } catch {
            Write-Log "Add-PrinterDriver $n : $($_.Exception.Message)" -Level INFO
        }
    }

    $have = @()
    foreach ($n in $wanted) {
        if (Get-PrinterDriver -Name $n -ErrorAction SilentlyContinue) { $have += $n }
    }
    if ($have.Count -gt 0) {
        Write-Status -Name "Drivers imprimante" -Status OK -Detail ($have -join ", ")
        return $true
    }
    if ($staged -gt 0) {
        Write-Status -Name "Drivers imprimante" -Status INFO -Detail "$staged INF traite(s) — noms non encore visibles" -NoCount
        return $true
    }
    Write-Status -Name "Drivers imprimante" -Status FAIL -Detail "echec install offline — Add-Printer tentera les fallbacks"
    return $false
}

function Add-ProvTcpIpPrinter {
    <# Ajoute une imprimante TCP/IP (port Standard TCP/IP + driver). #>
    param([hashtable]$Printer)
    $name = $Printer.Name
    $ip   = $Printer.HostAddress
    $port = if ($Printer.PortName) { $Printer.PortName } else { "IP_$ip" }
    $wantDriver = if ($Printer.DriverName) { $Printer.DriverName } else { "Microsoft IPP Class Driver" }

    if (-not $name -or -not $ip) {
        Write-Status -Name "Imprimante TCP" -Status FAIL -Detail "Name/HostAddress manquant"
        return $false
    }

    $existing = Get-Printer -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -eq $name -or $_.PortName -eq $port -or $_.PortName -eq $ip
    } | Select-Object -First 1
    if ($existing) {
        Write-Status -Name "Imprimante" -Status SKIP -Detail "$name ($($existing.PortName))" -NoCount
        return $false
    }

    try {
        $pExist = Get-PrinterPort -Name $port -ErrorAction SilentlyContinue
        if (-not $pExist) {
            Add-PrinterPort -Name $port -PrinterHostAddress $ip -ErrorAction Stop
            Write-Log "Port TCP/IP cree: $port -> $ip" -Level INFO
        }
    } catch {
        Write-Status -Name "Imprimante $name" -Status FAIL -Detail "port $port — $($_.Exception.Message)"
        return $false
    }

    $candidates = @(
        $wantDriver
    ) + @($ProvConfig.PrinterDriverNames | Where-Object { $_ }) + @(
        "Microsoft IPP Class Driver",
        "MS Publisher Imagesetter"
    ) | Select-Object -Unique

    $driverOk = $null
    foreach ($d in $candidates) {
        $hit = Get-PrinterDriver -Name $d -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hit) { $driverOk = $d; break }
        try {
            Add-PrinterDriver -Name $d -ErrorAction Stop
            $driverOk = $d
            break
        } catch {}
    }

    if (-not $driverOk) {
        Write-Status -Name "Imprimante $name" -Status FAIL -Detail "aucun driver (voulu: $wantDriver) — installez $(@($ProvConfig.PrinterDriverNames)[0]) puis relancez"
        return $false
    }

    try {
        Add-Printer -Name $name -DriverName $driverOk -PortName $port -ErrorAction Stop
        Write-Status -Name "Imprimante" -Status OK -Detail "$name @ $ip ($driverOk)"
        return $true
    } catch {
        Write-Status -Name "Imprimante $name" -Status FAIL -Detail $_.Exception.Message
        return $false
    }
}

function Find-ProvJavaW {
    $hit = @(
        "${env:ProgramFiles}\Java\jre*\bin\javaw.exe",
        "${env:ProgramFiles}\Java\jdk*\bin\javaw.exe",
        "${env:ProgramFiles}\Eclipse Adoptium\jre-*\bin\javaw.exe",
        "${env:ProgramFiles}\Eclipse Adoptium\jdk-*\bin\javaw.exe",
        "${env:ProgramFiles}\Microsoft\jdk-*\bin\javaw.exe",
        "${env:ProgramFiles(x86)}\Java\jre*\bin\javaw.exe"
    ) | ForEach-Object { Get-Item $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
    if ($hit) { return $hit.FullName }
    $cmd = Get-Command javaw.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $cmd2 = Get-Command java.exe -ErrorAction SilentlyContinue
    if ($cmd2) { return $cmd2.Source }
    return $null
}

function Find-ProvLobLauncherJar {
    param([string]$LobDir)
    if (-not $LobDir -or -not (Test-Path -LiteralPath $LobDir)) { return $null }
    $glob = "*launcher*.jar"
    if ($ProvConfig.LobApp -and $ProvConfig.LobApp.LauncherGlob) { $glob = [string]$ProvConfig.LobApp.LauncherGlob }
    $jar = Get-ChildItem -LiteralPath $LobDir -Filter $glob -File -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($jar) { return $jar }
    # Evite les jars bin\ — cherche a la racine de l'app metier seulement
    return Get-ChildItem -LiteralPath $LobDir -Filter "*.jar" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '(?i)launcher' } |
        Select-Object -First 1
}

function Stop-ProvLobSideProcesses {
    <# Ferme les process ouverts par le 1er lancement de l'app metier (config LobApp.KillAfterFirstRun). #>
    $killed = @()
    $names = @('javaw', 'java', 'chrome')
    if ($ProvConfig.LobApp -and $ProvConfig.LobApp.KillAfterFirstRun) { $names = @($ProvConfig.LobApp.KillAfterFirstRun) }
    foreach ($name in $names) {
        $procs = @(Get-Process -Name $name -ErrorAction SilentlyContinue)
        foreach ($p in $procs) {
            try {
                Stop-Process -Id $p.Id -Force -ErrorAction Stop
                $killed += "$name($($p.Id))"
            } catch {}
        }
    }
    return $killed
}

function Start-ProvLobLauncher {
    <#
    Lance le JAR launcher de l'app metier (raccourci Bureau + Chrome).
    Attend confirmation operateur, puis ferme les process listes avant de poursuivre.
    #>
    param([string]$LobDir)
    $lobLabel = [string]$ProvConfig.LobApp.Label
    if (-not $lobLabel) { $lobLabel = [string]$ProvConfig.LobApp.Name }
    $jar = Find-ProvLobLauncherJar -LobDir $LobDir
    if (-not $jar) {
        Write-Status -Name "$lobLabel Launcher" -Status FAIL -Detail "$($ProvConfig.LobApp.LauncherGlob) introuvable dans $LobDir"
        return $false
    }
    $java = Find-ProvJavaW
    if (-not $java) {
        Write-Status -Name "$lobLabel Launcher" -Status FAIL -Detail "javaw/java introuvable — installer Java puis relancer"
        return $false
    }
    try {
        Start-Process -FilePath $java -ArgumentList @("-jar", $jar.FullName) -WorkingDirectory $LobDir -WindowStyle Normal
        Write-Status -Name "$lobLabel Launcher" -Status OK -Detail "$($jar.Name) demarre"
        Write-Log "$lobLabel launcher started: $java -jar $($jar.FullName)" -Level INFO
    } catch {
        Write-Status -Name "$lobLabel Launcher" -Status FAIL -Detail $_.Exception.Message
        return $false
    }

    $boxTitle = " $lobLabel "
    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + $boxTitle + ($UI.HLine * [Math]::Max(4, 52 - $boxTitle.Length))) -ForegroundColor Magenta
    Write-Host ("  " + $UI.VLine + "  Java + Chrome vont s'ouvrir (raccourci Bureau).") -ForegroundColor Gray
    Write-Host ("  " + $UI.VLine + "  Quand c'est bon : O / Oui / Entree — on ferme Java+Chrome puis on continue.") -ForegroundColor Yellow
    Write-Host ("  " + $UI.VLine + "  N = pas encore pret (le script attend, relance possible).") -ForegroundColor DarkGray
    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
    Write-Host ""

    $ok = $false
    while (-not $ok) {
        $ok = Read-ProvYesNo -Prompt "$lobLabel OK ?" -Default Yes
        if (-not $ok) {
            Write-Host ("     OK — on attend. Relancer le launcher ?") -ForegroundColor DarkGray
            $relance = Read-ProvYesNo -Prompt "Relancer $lobLabel Launcher ?" -Default No
            if ($relance) {
                try {
                    Start-Process -FilePath $java -ArgumentList @("-jar", $jar.FullName) -WorkingDirectory $LobDir -WindowStyle Normal
                } catch {}
            }
        }
    }

    $killed = @(Stop-ProvLobSideProcesses)
    if ($killed.Count -gt 0) {
        Write-Status -Name "$lobLabel cleanup" -Status OK -Detail ("ferme : " + ($killed -join ", "))
    } else {
        Write-Status -Name "$lobLabel cleanup" -Status INFO -Detail "aucun process java/chrome a fermer" -NoCount
    }
    return $true
}

function Install-ProvLobApps {
    <#
    Apps metier hors winget (declarees dans config.local.psd1) :
      - PortableApp  : copie d'un dossier portable + raccourci menu Demarrer
      - LobApp       : copie dossier + lance le launcher JAR (raccourci Bureau)
      - ClickOnceApp : application ClickOnce (reseau requis)
      - Client camera : offline / URL archive / winget (si CameraClientEnabled)
    #>
    Write-Section -Title "Applications metier (offline)" -Step "4b/8"
    $roots = @(Get-ProvOfflineRoots)
    if ($roots.Count -eq 0) {
        Write-Status -Name "Sources offline" -Status FAIL -Detail "aucune (brancher la cle USB avec scripts\offline)" -NoCount
    } else {
        Write-Status -Name "Sources offline" -Status INFO -Detail ($roots -join "  |  ") -NoCount
        Write-Log ("Offline roots: " + ($roots -join "; ")) -Level INFO
    }
    $destRoot = [string]$ProvConfig.LocalAppRoot
    if (-not (Test-Path $destRoot)) {
        New-Item -ItemType Directory -Path $destRoot -Force | Out-Null
    }

    # --- Application portable (copie dossier) ---
    $pa = $ProvConfig.PortableApp
    if ($pa -and $pa.Folder) {
        $paSrc = Find-ProvOfflineItem -RelativeNames @($pa.Folder) -Directory
        $paDst = Join-Path $destRoot $pa.Folder
        if ($paSrc) {
            $paAlready = $false
            if (Test-Path -LiteralPath (Join-Path $paDst $pa.Exe)) {
                $paAlready = $true
            } elseif (Get-ChildItem -LiteralPath $paDst -Filter $pa.Exe -Recurse -EA SilentlyContinue | Select-Object -First 1) {
                $paAlready = $true
            }
            if ($paAlready) {
                Write-Status -Name $pa.Name -Status SKIP -Detail $paDst
            } else {
                Write-Status -Name $pa.Name -Status RUN -Detail "copie vers $paDst..." -NoCount
                try {
                    [void](Copy-ProvDirectoryRobust -Source $paSrc -Destination $paDst)
                    $exe = Get-ChildItem -LiteralPath $paDst -Filter $pa.Exe -Recurse -EA SilentlyContinue | Select-Object -First 1
                    if ($exe) {
                        New-ProvStartMenuShortcut -Name $pa.Name -TargetPath $exe.FullName | Out-Null
                        Write-Status -Name $pa.Name -Status OK -Detail $exe.FullName
                    } else {
                        Write-Status -Name $pa.Name -Status FAIL -Detail "copie OK mais exe introuvable"
                    }
                } catch {
                    Write-Status -Name $pa.Name -Status FAIL -Detail $_.Exception.Message
                }
            }
        } else {
            Write-Status -Name $pa.Name -Status FAIL -Detail "dossier offline\$($pa.Folder) manquant"
        }
    }

    # --- App metier (dossier + launcher JAR) ---
    $lob = $ProvConfig.LobApp
    if ($lob -and $lob.SourceFolders) {
        $lobSrc = Find-ProvOfflineItem -RelativeNames @($lob.SourceFolders) -Directory
        $lobDst = Join-Path $destRoot $lob.DestFolder
        $lobReady = $false
        if ($lobSrc) {
            $jarExisting = Find-ProvLobLauncherJar -LobDir $lobDst
            if ($jarExisting) {
                Write-Status -Name $lob.Name -Status SKIP -Detail $lobDst
                $lobReady = $true
            } else {
                Write-Status -Name $lob.Name -Status RUN -Detail "copie vers $lobDst..." -NoCount
                try {
                    [void](Copy-ProvDirectoryRobust -Source $lobSrc -Destination $lobDst)
                    $jar = Find-ProvLobLauncherJar -LobDir $lobDst
                    $javaPath = Find-ProvJavaW
                    if ($jar -and $javaPath) {
                        $ico = Get-ChildItem -LiteralPath $lobDst -Filter "*.ico" -EA SilentlyContinue | Select-Object -First 1
                        $icoPath = if ($ico) { $ico.FullName } else { "" }
                        New-ProvStartMenuShortcut -Name $lob.Shortcut -TargetPath $javaPath `
                            -Arguments ("-jar `"$($jar.FullName)`"") -WorkingDirectory $lobDst -IconLocation $icoPath | Out-Null
                        Write-Status -Name $lob.Name -Status OK -Detail $jar.Name
                        $lobReady = $true
                    } elseif ($jar) {
                        New-ProvStartMenuShortcut -Name $lob.Shortcut -TargetPath $jar.FullName -WorkingDirectory $lobDst | Out-Null
                        Write-Status -Name $lob.Name -Status OK -Detail "$($jar.Name) (assoc. .jar)"
                        $lobReady = $true
                    } else {
                        Write-Status -Name $lob.Name -Status FAIL -Detail "copie OK mais jar introuvable"
                    }
                } catch {
                    Write-Status -Name $lob.Name -Status FAIL -Detail $_.Exception.Message
                }
            }
        } else {
            Write-Status -Name $lob.Name -Status FAIL -Detail "dossier offline\$(@($lob.SourceFolders)[0]) manquant"
        }

        # Lance le Launcher apres copie (ou si deja present) -> raccourci Bureau
        if ($lobReady) {
            Start-ProvLobLauncher -LobDir $lobDst | Out-Null
        }
    }

    # --- Application ClickOnce ---
    $co = $ProvConfig.ClickOnceApp
    if ($co -and $co.Name) {
        $coApp = $null
        if ($co.AppRefNames) { $coApp = Find-ProvOfflineItem -RelativeNames @($co.AppRefNames) }
        $coUrlFile = $null
        if ($co.UrlFileName) { $coUrlFile = Find-ProvOfflineItem -RelativeNames @($co.UrlFileName) }
        $coUrl = [string]$co.Url
        if ($coUrlFile) {
            try { $coUrl = (Get-Content -LiteralPath $coUrlFile -Raw -Encoding UTF8).Trim().Split("#")[0].Trim() } catch {}
        } elseif ($coApp) {
            try {
                $raw = [System.IO.File]::ReadAllText($coApp, [System.Text.Encoding]::Unicode).Trim()
                if ($raw -match '^https?://') { $coUrl = $raw.Split("#")[0].Trim() }
            } catch {}
        }

        $coInstalled = $null
        if ($co.InstalledFilter -and $co.InstalledPattern) {
            $coInstalled = Get-ChildItem "$env:LOCALAPPDATA\Apps\2.0" -Filter $co.InstalledFilter -Recurse -EA SilentlyContinue |
                Where-Object { $_.Name -match $co.InstalledPattern } | Select-Object -First 1
        }
        $coLnk = $null
        if ($co.ShortcutNames) { $coLnk = Find-Shortcut -Names @($co.ShortcutNames) }

        if ($coInstalled -or $coLnk) {
            Write-Status -Name $co.Name -Status SKIP -Detail "deja present"
        } else {
            Write-Status -Name $co.Name -Status RUN -Detail "ClickOnce..." -NoCount
            try {
                if ($coApp) {
                    $menu = Join-Path $env:ProgramData (Join-Path "Microsoft\Windows\Start Menu\Programs" $ProvConfig.StartMenuFolder)
                    if (-not (Test-Path $menu)) { New-Item -ItemType Directory -Path $menu -Force | Out-Null }
                    Copy-Item -LiteralPath $coApp -Destination (Join-Path $menu (@($co.AppRefNames)[0])) -Force
                }
                # Declenche l'install ClickOnce (necessite reseau)
                $p = Start-Process -FilePath "$env:SystemRoot\System32\rundll32.exe" `
                    -ArgumentList "dfshim.dll,ShOpenVerbApplication $coUrl" -PassThru -WindowStyle Normal -ErrorAction Stop
                Start-Sleep -Seconds 8
                Write-Status -Name $co.Name -Status OK -Detail "ClickOnce lance (valider si invite UAC/reseau)"
                Write-Log "ClickOnce $($co.Name) URL=$coUrl pid=$($p.Id)" -Level INFO
            } catch {
                Write-Status -Name $co.Name -Status FAIL -Detail $_.Exception.Message
            }
        }
    }

    # --- Client camera (offline -> URL archive -> winget) ---
    if ($ProvConfig.CameraClientEnabled) {
        Install-ProvCameraClient
    }
}

function Get-ProvCameraClientName {
    $n = [string]$ProvConfig.CameraClientName
    if (-not $n) { $n = "Client camera" }
    return $n
}

function Find-ProvCameraClientExe {
    <# Cherche l'exe du client camera : chemins connus (config) puis dossiers Program Files correspondants. #>
    $paths = @()
    foreach ($rel in @($ProvConfig.CameraClientInstalledPaths | Where-Object { $_ })) {
        $paths += (Join-Path ${env:ProgramFiles} $rel)
        if (${env:ProgramFiles(x86)}) { $paths += (Join-Path ${env:ProgramFiles(x86)} $rel) }
    }
    $hit = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($hit) { return $hit }
    if (-not $ProvConfig.CameraClientProgramDirPattern -or -not $ProvConfig.CameraClientExeNamePattern) { return $null }
    foreach ($root in @(${env:ProgramFiles}, ${env:ProgramFiles(x86)})) {
        if (-not $root -or -not (Test-Path $root)) { continue }
        $found = Get-ChildItem -LiteralPath $root -Directory -EA SilentlyContinue |
            Where-Object { $_.Name -match $ProvConfig.CameraClientProgramDirPattern } |
            ForEach-Object {
                Get-ChildItem -LiteralPath $_.FullName -Recurse -Filter "*.exe" -EA SilentlyContinue |
                    Where-Object { $_.Name -match $ProvConfig.CameraClientExeNamePattern } | Select-Object -First 1
            } | Select-Object -First 1
        if ($found) { return $found.FullName }
    }
    return $null
}

function Install-ProvCameraClientSetupFile {
    param([string]$SetupPath)
    if (-not (Test-Path -LiteralPath $SetupPath)) { return $false }
    $ccName = Get-ProvCameraClientName
    $ext = [IO.Path]::GetExtension($SetupPath).ToLowerInvariant()
    Write-Status -Name $ccName -Status RUN -Detail "install $(Split-Path $SetupPath -Leaf)..." -NoCount
    try {
        if ($ext -eq ".msi") {
            $p = Start-Process msiexec.exe -ArgumentList "/i `"$SetupPath`" /qn /norestart" -Wait -PassThru -NoNewWindow
            Write-Log "$ccName MSI exit=$($p.ExitCode)" -Level INFO
        } else {
            # InstallShield / NSIS : essais silencieux courants
            $argSets = @(
                @("/S"),
                @("/silent"),
                @("/quiet"),
                @("/s", "/v`"/qn /norestart`"")
            )
            $p = $null
            foreach ($setupArgs in $argSets) {
                $p = Start-Process -FilePath $SetupPath -ArgumentList $setupArgs -Wait -PassThru -ErrorAction SilentlyContinue
                if (Find-ProvCameraClientExe) { break }
            }
            if ($p) { Write-Log "$ccName setup exit=$($p.ExitCode)" -Level INFO }
        }
    } catch {
        Write-Log "$ccName setup fail: $($_.Exception.Message)" -Level WARN
        return $false
    }
    return [bool](Find-ProvCameraClientExe)
}

function Install-ProvCameraClient {
    <#
    Client de videosurveillance (optionnel, config CameraClient*). Ordre :
      1) deja installe
      2) offline\ (exe/msi/zip correspondant a CameraClientFilePattern)
      3) telechargement CameraClientArchiveUrl (archive zip)
      4) winget CameraClientWingetId (paquet winget)
    #>
    $ccName = Get-ProvCameraClientName
    $ccPkg = [string]$ProvConfig.CameraClientWingetName
    if (-not $ccPkg) { $ccPkg = "$ccName (winget)" }
    $present = Find-ProvCameraClientExe
    if ($present) {
        Write-Status -Name $ccName -Status SKIP -Detail $present
        return
    }

    # --- 2) offline local ---
    $setup = $null
    $zipLocal = $null
    foreach ($root in @(Get-ProvOfflineRoots)) {
        if ($ProvConfig.CameraClientFilePattern) {
            $cands = Get-ChildItem -LiteralPath $root -File -EA SilentlyContinue |
                Where-Object { $_.Name -match $ProvConfig.CameraClientFilePattern -and $_.Extension -match '\.(exe|msi|zip)$' }
            foreach ($c in $cands) {
                if ($c.Extension -eq ".zip") { if (-not $zipLocal) { $zipLocal = $c } }
                else { $setup = $c; break }
            }
            if ($setup) { break }
        }
        if ($ProvConfig.CameraClientOfflineSubdir) {
            $sub = Join-Path $root $ProvConfig.CameraClientOfflineSubdir
            if (Test-Path $sub) {
                $cands = Get-ChildItem -LiteralPath $sub -Recurse -Include *.exe,*.msi -EA SilentlyContinue |
                    Where-Object { $_.Name -notmatch '(?i)unins' } | Select-Object -First 1
                if ($cands) { $setup = $cands; break }
            }
        }
    }

    if ($setup) {
        if (Install-ProvCameraClientSetupFile -SetupPath $setup.FullName) {
            $exe = Find-ProvCameraClientExe
            New-ProvStartMenuShortcut -Name $ccName -TargetPath $exe | Out-Null
            Write-Status -Name $ccName -Status OK -Detail "offline: $exe"
            return
        }
        Write-Status -Name $ccName -Status INFO -Detail "setup offline echoue — suite DL / winget" -NoCount
    }

    # Extraire un zip offline si present
    $work = Join-Path $env:TEMP ("{0}-{1}" -f $ProvConfig.WorkFolderName, ($ccName -replace '[^\w\-]', ''))
    if (-not (Test-Path $work)) { New-Item -ItemType Directory -Path $work -Force | Out-Null }
    $innerPattern = [string]$ProvConfig.CameraClientExePattern
    if (-not $innerPattern) { $innerPattern = '.' }

    if ($zipLocal -and -not (Find-ProvCameraClientExe)) {
        Write-Status -Name $ccName -Status RUN -Detail "extraction zip offline..." -NoCount
        try {
            $extract = Join-Path $work "offline-extract"
            if (Test-Path $extract) { Remove-Item $extract -Recurse -Force -EA SilentlyContinue }
            Expand-Archive -LiteralPath $zipLocal.FullName -DestinationPath $extract -Force
            $inner = Get-ChildItem -LiteralPath $extract -Recurse -Include *.exe,*.msi -EA SilentlyContinue |
                Where-Object { $_.Name -match $innerPattern -and $_.Name -notmatch '(?i)unins' } |
                Select-Object -First 1
            if ($inner -and (Install-ProvCameraClientSetupFile -SetupPath $inner.FullName)) {
                $exe = Find-ProvCameraClientExe
                New-ProvStartMenuShortcut -Name $ccName -TargetPath $exe | Out-Null
                Write-Status -Name $ccName -Status OK -Detail "zip offline: $exe"
                return
            }
        } catch {
            Write-Log "$ccName zip offline: $($_.Exception.Message)" -Level WARN
        }
    }

    # --- 3) URL archive ---
    $url = [string]$ProvConfig.CameraClientArchiveUrl
    $ver = [string]$ProvConfig.CameraClientArchiveVersion
    if ($url -and -not (Find-ProvCameraClientExe)) {
        Write-Status -Name $ccName -Status RUN -Detail "DL archive $ver..." -NoCount
        $got = $false
        $zipPath = $null
        # Cache zip : TEMP, ou offline USB deja present (ne jamais creer offline\ sur le Bureau)
        if ($ProvConfig.CameraClientCacheZipFormat) {
            foreach ($root in @(Get-ProvOfflineRoots)) {
                if (-not (Test-ProvPathWritable -Dir $root)) { continue }
                $cacheZip = Join-Path $root ([string]$ProvConfig.CameraClientCacheZipFormat -f $ver)
                if ((Test-Path $cacheZip) -and ((Get-Item $cacheZip).Length -gt 1MB)) {
                    $zipPath = $cacheZip
                    $got = $true
                    Write-Status -Name $ccName -Status INFO -Detail "cache offline utilise" -NoCount
                    break
                }
            }
        }
        if (-not $got) {
            $dlName = if ($ProvConfig.CameraClientDownloadZipFormat) { [string]$ProvConfig.CameraClientDownloadZipFormat -f $ver } else { "camera-client-$ver.zip" }
            $zipPath = Join-Path $work $dlName
            $got = Save-ProvDownload -Url $url -OutFile $zipPath
        }

        if ($got -and (Test-Path $zipPath) -and ((Get-Item $zipPath).Length -gt 1MB)) {
            try {
                $extract = Join-Path $work "archive-extract"
                if (Test-Path $extract) { Remove-Item $extract -Recurse -Force -EA SilentlyContinue }
                Write-Status -Name $ccName -Status RUN -Detail "extraction + install silencieuse..." -NoCount
                Expand-Archive -LiteralPath $zipPath -DestinationPath $extract -Force
                $inner = Get-ChildItem -LiteralPath $extract -Recurse -Include *.exe,*.msi -EA SilentlyContinue |
                    Where-Object { $_.Name -match $innerPattern -and $_.Name -notmatch '(?i)unins' } |
                    Sort-Object Length -Descending | Select-Object -First 1
                if (-not $inner) {
                    Write-Status -Name $ccName -Status FAIL -Detail "pas d'exe dans le zip telecharge"
                } elseif (Install-ProvCameraClientSetupFile -SetupPath $inner.FullName) {
                    $exe = Find-ProvCameraClientExe
                    New-ProvStartMenuShortcut -Name $ccName -TargetPath $exe | Out-Null
                    Write-Status -Name $ccName -Status OK -Detail "archive $ver"
                    return
                } else {
                    Write-Status -Name $ccName -Status INFO -Detail "install archive non silencieuse / incomplete" -NoCount
                }
            } catch {
                Write-Status -Name $ccName -Status FAIL -Detail "zip telecharge: $($_.Exception.Message)"
            }
        } else {
            Write-Status -Name $ccName -Status INFO -Detail "DL archive impossible (reseau) — fallback winget" -NoCount
        }
    }

    # --- 4) winget (paquet) ---
    if (Find-ProvCameraClientExe) { return }
    if (-not $ProvConfig.CameraClientWingetId) {
        Write-Status -Name $ccName -Status FAIL -Detail "offline + URL echoues — placez le zip/exe dans offline\"
        return
    }
    Write-Status -Name $ccPkg -Status INFO -Detail "fallback winget" -NoCount
    $st = Install-WingetApp -Name $ccPkg -Id ([string]$ProvConfig.CameraClientWingetId) -Retries 1
    $exe = Find-ProvCameraClientExe
    if ($exe) {
        New-ProvStartMenuShortcut -Name $ccPkg -TargetPath $exe | Out-Null
        Write-Status -Name $ccPkg -Status OK -Detail $exe
        return
    }
    if ($st -eq "OK" -or $st -eq "SKIP") {
        Write-Status -Name $ccPkg -Status OK -Detail "winget $st (chemin exe non detecte)"
        return
    }
    Write-Status -Name $ccName -Status FAIL -Detail "URL + winget echoues — placez le zip/exe dans offline\"
}

function Set-ProvNetworkResources {
    Write-Status -Name "Ressources reseau" -Status RUN -Detail "lecteurs + imprimantes" -NoCount

    # Drivers imprimante depuis offline\printers AVANT Add-Printer
    Install-ProvOfflinePrinterDrivers | Out-Null

    $driveCount = 0
    foreach ($d in @($ProvConfig.NetworkDrives)) {
        if (-not $d -or -not $d.Path) { continue }
        $letter = ($d.Letter.ToString().TrimEnd(':') + ":")
        try {
            if (Test-Path $letter) {
                Write-Status -Name "Lecteur $letter" -Status SKIP -Detail $d.Path -NoCount
            } else {
                New-PSDrive -Name $d.Letter.ToString().TrimEnd(':') -PSProvider FileSystem -Root $d.Path -Persist -ErrorAction Stop | Out-Null
                # Persist via net use aussi (plus fiable au logon)
                & net.exe use $letter $d.Path /persistent:yes 1>$null 2>$null
                Write-Status -Name "Lecteur $letter" -Status OK -Detail $d.Path
                $driveCount++
            }
        } catch {
            Write-Status -Name "Lecteur $letter" -Status FAIL -Detail $_.Exception.Message
        }
    }

    $printerCount = 0
    foreach ($pr in @($ProvConfig.Printers)) {
        if (-not $pr) { continue }
        if ($pr.HostAddress) {
            if (Add-ProvTcpIpPrinter -Printer $pr) { $printerCount++ }
            continue
        }
        if (-not $pr.Path) { continue }
        try {
            $exists = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $pr.Name -or $_.PortName -eq $pr.Path }
            if ($exists) {
                Write-Status -Name "Imprimante" -Status SKIP -Detail $pr.Name -NoCount
            } else {
                Add-Printer -ConnectionName $pr.Path -ErrorAction Stop
                Write-Status -Name "Imprimante" -Status OK -Detail $pr.Path
                $printerCount++
            }
        } catch {
            Write-Status -Name "Imprimante" -Status FAIL -Detail "$($pr.Path) — $($_.Exception.Message)"
        }
    }

    # Prompt libre optionnel
    Write-Host ""
    $extra = (Read-Host "     Lecteur UNC supplementaire [Entree = aucun] ex. \\\\serveur\\partage").Trim()
    if ($extra) {
        $let = (Read-Host "     Lettre (ex. S)").Trim().TrimEnd(':')
        if ($let) {
            try {
                & net.exe use "${let}:" $extra /persistent:yes 1>$null 2>$null
                if ($LASTEXITCODE -eq 0) {
                    Write-Status -Name "Lecteur ${let}:" -Status OK -Detail $extra
                    $driveCount++
                } else {
                    Write-Status -Name "Lecteur ${let}:" -Status FAIL -Detail "net use exit $LASTEXITCODE"
                }
            } catch {
                Write-Status -Name "Lecteur ${let}:" -Status FAIL -Detail $_.Exception.Message
            }
        }
    }

    if ($driveCount -eq 0 -and $printerCount -eq 0 -and $ProvConfig.NetworkDrives.Count -eq 0 -and $ProvConfig.Printers.Count -eq 0) {
        Write-Status -Name "Ressources reseau" -Status INFO -Detail "aucune config (NetworkDrives / Printers dans config.local.psd1)" -NoCount
    } else {
        Write-Status -Name "Ressources reseau" -Status OK -Detail "$driveCount lecteur(s), $printerCount imprimante(s)"
    }
}

function Set-ProvAnyDeskPassword {
    <#
    Mot de passe d'acces non supervise AnyDesk.
    Ce n'est PAS un mot de passe a "recuperer" quelque part : VOUS le choisissez ici.
    Il servira pour vous connecter a distance sans validation a l'ecran.
    #>
    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " AnyDesk — acces distant " + ($UI.HLine * 28)) -ForegroundColor Magenta
    Write-Host ("  " + $UI.VLine + "  Le mot de passe est CHOISI par vous (pas recupere).") -ForegroundColor Gray
    Write-Host ("  " + $UI.VLine + "  Notez-le (Bitwarden) : acces unattended au poste.") -ForegroundColor DarkGray
    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
    Write-Host ""

    $any = @(
        "${env:ProgramFiles(x86)}\AnyDesk\AnyDesk.exe",
        "${env:ProgramFiles}\AnyDesk\AnyDesk.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $any) {
        Write-Status -Name "AnyDesk password" -Status FAIL -Detail "AnyDesk.exe introuvable"
        return
    }

    # Afficher l'ID AnyDesk si possible
    try {
        $idFile = Join-Path $env:APPDATA "AnyDesk\system.conf"
        if (-not (Test-Path $idFile)) { $idFile = Join-Path ${env:ProgramData} "AnyDesk\system.conf" }
        if (Test-Path $idFile) {
            $ad = Select-String -Path $idFile -Pattern "ad\.anynet\.id\s*=\s*(.+)" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($ad) {
                $script:AnyDeskId = $ad.Matches.Groups[1].Value.Trim()
                Write-Host ("     AnyDesk ID : $($script:AnyDeskId)") -ForegroundColor Cyan
            }
        }
    } catch {}

    $setPwd = (Read-Host "     Definir un mot de passe unattended ? [O/n]").Trim().ToLowerInvariant()
    if ($setPwd -eq "n" -or $setPwd -eq "non") {
        Write-Status -Name "AnyDesk password" -Status SKIP -Detail "ignore"
        return
    }

    $pwd = Read-Host "     Mot de passe AnyDesk unattended" -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwd)
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) | Out-Null
    if ([string]::IsNullOrWhiteSpace($plain)) {
        Write-Status -Name "AnyDesk password" -Status SKIP -Detail "vide"
        return
    }

    try {
        # Documentation AnyDesk : echo pwd | AnyDesk.exe --set-password
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $any
        $psi.Arguments = "--set-password"
        $psi.RedirectStandardInput = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $proc = [Diagnostics.Process]::Start($psi)
        $proc.StandardInput.WriteLine($plain)
        $proc.StandardInput.Close()
        $proc.WaitForExit(15000) | Out-Null
        Write-Status -Name "AnyDesk password" -Status OK -Detail "mot de passe unattended defini"
        Write-Log "AnyDesk unattended password defini (non journalise en clair)" -Level INFO
    } catch {
        Write-Status -Name "AnyDesk password" -Status FAIL -Detail $_.Exception.Message
    }
}

function Remove-ProvBloatware {
    Write-Status -Name "Bloat Win11" -Status RUN -Detail "desinstallation apps inutiles" -NoCount
    $packages = @(
        "Microsoft.BingNews",
        "Microsoft.BingWeather",
        "Microsoft.BingSearch",
        "Microsoft.GetHelp",
        "Microsoft.Getstarted",
        "Microsoft.MicrosoftOfficeHub",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.People",
        "Microsoft.PowerAutomateDesktop",
        "Microsoft.Todos",
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.WindowsMaps",
        "Microsoft.Xbox.TCUI",
        "Microsoft.XboxGameOverlay",
        "Microsoft.XboxGamingOverlay",
        "Microsoft.XboxIdentityProvider",
        "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.GamingApp",
        "Microsoft.ZuneMusic",
        "Microsoft.ZuneVideo",
        "Microsoft.Clipchamp",
        "Microsoft.YourPhone",
        "MicrosoftTeams",
        "Microsoft.MicrosoftStickyNotes",
        "king.com.CandyCrushSaga",
        "king.com.CandyCrushSodaSaga",
        "Disney.37853FC22B2CE",
        "Facebook.Facebook*",
        "SpotifyAB.SpotifyMusic"
    )
    $removed = 0
    foreach ($pat in $packages) {
        $apps = Get-AppxPackage -Name $pat -AllUsers -ErrorAction SilentlyContinue
        if (-not $apps) { $apps = Get-AppxPackage -Name $pat -ErrorAction SilentlyContinue }
        foreach ($app in @($apps)) {
            try {
                Remove-AppxPackage -Package $app.PackageFullName -ErrorAction Stop
                $removed++
                Write-Log "Bloat retire : $($app.Name)" -Level INFO
            } catch {
                try {
                    Remove-AppxPackage -Package $app.PackageFullName -AllUsers -ErrorAction Stop
                    $removed++
                } catch {
                    Write-Log "Bloat KO $($app.Name): $($_.Exception.Message)" -Level WARN
                }
            }
        }
        # Provisioned (nouveaux profils)
        Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like $pat } |
            ForEach-Object {
                try {
                    Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue | Out-Null
                } catch {}
            }
    }
    Write-Status -Name "Bloat Win11" -Status OK -Detail "$removed package(s) retire(s)"
}

function Start-ProvWindowsUpdate {
    Write-Status -Name "Windows Update" -Status RUN -Detail "scan + telechargement..." -NoCount
    try {
        # Declenche le service Update Orchestrator (Win10/11)
        $uso = Join-Path $env:SystemRoot "System32\UsoClient.exe"
        if (Test-Path $uso) {
            Start-Process $uso -ArgumentList "StartScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
            Start-Process $uso -ArgumentList "StartDownload" -WindowStyle Hidden -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Start-Process $uso -ArgumentList "StartInstall" -WindowStyle Hidden -ErrorAction SilentlyContinue
            Write-Status -Name "Windows Update" -Status OK -Detail "scan/download/install declenches (arriere-plan)"
        } else {
            Write-Status -Name "Windows Update" -Status FAIL -Detail "UsoClient.exe introuvable"
        }
    } catch {
        Write-Status -Name "Windows Update" -Status FAIL -Detail $_.Exception.Message
    }
}

function Test-ProvIntuneCompliance {
    Write-Status -Name "Intune / Entra ID" -Status RUN -Detail "dsregcmd /status" -NoCount
    try {
        $status = & dsregcmd.exe /status 2>$null | Out-String
        Write-Log ($status.Substring(0, [Math]::Min(2000, $status.Length))) -Level INFO

        $azureAd = if ($status -match "AzureAdJoined\s*:\s*YES") { "YES" } else { "NO" }
        $domain  = if ($status -match "DomainJoined\s*:\s*YES") { "YES" } else { "NO" }
        $prt     = if ($status -match "AzureAdPrt\s*:\s*YES") { "YES" } else { "NO" }
        $mdm     = if ($status -match "MDMUrl\s*:\s*http") { "YES" } else { "NO" }

        $detail = "AzureAD=$azureAd PRT=$prt MDM=$mdm Domain=$domain"
        if ($azureAd -eq "YES" -and $prt -eq "YES") {
            Write-Status -Name "Intune / Entra ID" -Status OK -Detail $detail
        } elseif ($azureAd -eq "YES") {
            Write-Status -Name "Intune / Entra ID" -Status SKIP -Detail "$detail (PRT absent = reconnexion compte)"
        } else {
            Write-Status -Name "Intune / Entra ID" -Status FAIL -Detail "$detail — joindre Entra ID / Company Portal"
        }
        return [pscustomobject]@{ AzureAdJoined = $azureAd; PRT = $prt; MDM = $mdm; DomainJoined = $domain }
    } catch {
        Write-Status -Name "Intune / Entra ID" -Status FAIL -Detail $_.Exception.Message
        return $null
    }
}

function Export-ProvInventory {
    $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostName = $script:ProvComputerName
    if (-not $hostName) { $hostName = $env:COMPUTERNAME }

    $rows = @()
    # winget list
    try {
        $wg = & winget list --accept-source-agreements --disable-interactivity 2>$null | Out-String
        foreach ($line in ($wg -split "`r?`n")) {
            if ($line -match '^\s*$' -or $line -match '^Name\s+Id' -or $line -match '^-+' -or $line -match '^$') { continue }
            $rows += [pscustomobject]@{ Source = "winget"; Line = $line.Trim() }
        }
    } catch {}

    # Apps registre
    $uninst = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    Get-ItemProperty $uninst -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName } |
        ForEach-Object {
            $rows += [pscustomobject]@{
                Source  = "registry"
                Line    = ("{0} | {1}" -f $_.DisplayName, $_.DisplayVersion)
            }
        }

    $csvName = "inventaire-$hostName-$stamp.csv"
    $txtName = "rapport-$hostName-$stamp.txt"

    # Documents\<ReportFolderName> (jamais le Bureau)
    $outDir = $script:ProvReportDir
    if (-not $outDir) { $outDir = Get-ProvReportDir; $script:ProvReportDir = $outDir }
    if (-not $outDir -or -not (Test-ProvPathWritable -Dir $outDir -Create)) {
        Write-Status -Name "Inventaire" -Status FAIL -Detail "impossible de creer Documents\$($ProvConfig.ReportFolderName)"
        return $null
    }

    $csvPath = Join-Path $outDir $csvName
    $txtPath = Join-Path $outDir $txtName

    try {
        $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    } catch {
        Write-Log "CSV inventaire KO: $($_.Exception.Message)" -Level WARN
        $csvPath = $null
    }

    $joinInfo = "n/a"
    try {
        $status = & dsregcmd.exe /status 2>$null | Out-String
        $azureAd = if ($status -match "AzureAdJoined\s*:\s*YES") { "YES" } else { "NO" }
        $prt     = if ($status -match "AzureAdPrt\s*:\s*YES") { "YES" } else { "NO" }
        $mdm     = if ($status -match "MDMUrl\s*:\s*http") { "YES" } else { "NO" }
        $joinInfo = "AzureAD=$azureAd PRT=$prt MDM=$mdm"
    } catch {}

    $report = @"
============================================================
$($ProvConfig.OrgName) — Rapport de deploiement
Date        : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
PC (actuel) : $env:COMPUTERNAME
PC (cible)  : $hostName
Utilisateur : $($script:ProvPrenom) $($script:ProvNom)
Email       : $($script:ProvEmail)
AnyDesk ID  : $($script:AnyDeskId)
------------------------------------------------------------
Stats       : OK=$($script:Stats.OK) DEJA=$($script:Stats.Skip) ECHEC=$($script:Stats.Fail)
Journal     : $LogFile
Inventaire  : $csvPath
Dossier     : $outDir
Entra/Intune: $joinInfo
============================================================
"@
    try {
        Set-Content -Path $txtPath -Value $report -Encoding UTF8
        $script:ReportPaths = @($txtPath)
        if ($csvPath) { $script:ReportPaths += $csvPath }
        if ($LogFile -and (Test-Path -LiteralPath $LogFile)) { $script:ReportPaths += $LogFile }
        Write-Status -Name "Inventaire / rapport" -Status OK -Detail $outDir
        return [pscustomobject]@{
            Txt      = $txtPath
            Csv      = $csvPath
            Log      = $LogFile
            Report   = $report
            OutDir   = $outDir
        }
    } catch {
        Write-Status -Name "Inventaire / rapport" -Status FAIL -Detail $_.Exception.Message
        return $null
    }
}

function Get-ProvReportAttachments {
    param($Export)
    $files = @()
    foreach ($p in @($Export.Txt, $Export.Csv, $Export.Log)) {
        if ($p -and (Test-Path -LiteralPath $p)) { $files += $p }
    }
    if ($script:ReportPaths) {
        foreach ($p in $script:ReportPaths) {
            if ($p -and (Test-Path -LiteralPath $p) -and ($files -notcontains $p)) { $files += $p }
        }
    }
    return $files
}

function Start-ProvMailtoUnelevated {
    <#
      mailto: depuis une session *admin* ne remonte souvent pas au client mail user (UAC).
      explorer.exe / rundll32 relancent le handler a integrite medium.
    #>
    param([Parameter(Mandatory)][string]$MailtoUri)
    if ([string]::IsNullOrWhiteSpace($MailtoUri)) { return $false }

    # 1) explorer = drop elevation (cas le plus fiable Win10/11)
    try {
        Start-Process -FilePath "$env:SystemRoot\explorer.exe" -ArgumentList $MailtoUri -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Write-Log "mailto explorer: $($_.Exception.Message)" -Level WARN
    }

    # 2) Protocol handler shell
    try {
        Start-Process -FilePath "$env:SystemRoot\System32\rundll32.exe" `
            -ArgumentList @("url.dll,FileProtocolHandler", $MailtoUri) -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Write-Log "mailto rundll32: $($_.Exception.Message)" -Level WARN
    }

    # 3) cmd start (dernier recours)
    try {
        Start-Process -FilePath "$env:SystemRoot\System32\cmd.exe" `
            -ArgumentList @("/c", "start", "", $MailtoUri) -WindowStyle Hidden -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Write-Log "mailto cmd: $($_.Exception.Message)" -Level WARN
    }

    return $false
}

function Send-ProvReport {
    param($Export)
    if (-not $Export) { return }

    if ($ProvConfig.IsExample) {
        # config.example.psd1 = destinataires fictifs : pas de mailto
        Write-Status -Name "Envoi email" -Status SKIP -Detail "config.example.psd1 — creer config.local.psd1 pour activer l'envoi"
        Write-Host ("     Rapport : $($Export.Txt)") -ForegroundColor Yellow
        return
    }
    $recipients = @(Get-ProvReportRecipients)
    if ($recipients.Count -eq 0) {
        $manual = (Read-Host "     Email destinataire du rapport [Entree = ignorer]").Trim()
        if ($manual) { $recipients = @($manual) }
    }
    if ($recipients.Count -eq 0) {
        Write-Status -Name "Envoi email" -Status SKIP -Detail "pas de destinataire"
        Write-Host ("     Rapport : $($Export.Txt)") -ForegroundColor Yellow
        return
    }
    $to = ($recipients -join ",")

    $subject = "$($ProvConfig.OrgName) deploiement — $($script:ProvComputerName) — $(Get-Date -Format 'yyyy-MM-dd')"
    $body = $Export.Report
    if ($script:ProvUsbReportDir) {
        $body = $body + "`r`nUSB miroir : $($script:ProvUsbReportDir)"
    }
    $attachments = @(Get-ProvReportAttachments -Export $Export)

    # Pas de SMTP / pas d'automatisation Outlook (evite setup compte Outlook 2016).
    # mailto multi-destinataires + dossier rapports pour glisser les PJ.
    try {
        if ($Export.OutDir -and (Test-Path $Export.OutDir)) {
            Start-Process explorer.exe -ArgumentList $Export.OutDir
        }
        $subjEnc = [uri]::EscapeDataString($subject)
        # Limiter le body : certains clients / Start-Process truncatent les longs mailto
        $bodyEnc = [uri]::EscapeDataString($body.Substring(0, [Math]::Min(900, $body.Length)))
        $mailto = "mailto:${to}?subject=${subjEnc}&body=${bodyEnc}"

        $opened = Start-ProvMailtoUnelevated -MailtoUri $mailto
        if ($opened) {
            Write-Status -Name "Envoi email" -Status OK -Detail ("mailto -> " + ($recipients -join ", "))
        } else {
            # Fichier .url cliquable (si aucun handler / elevation bloque encore)
            $urlFile = $null
            if ($Export.OutDir -and (Test-Path $Export.OutDir)) {
                $urlFile = Join-Path $Export.OutDir ("{0}-envoyer-rapport.url" -f $ProvConfig.OrgName)
            } else {
                $urlFile = Join-Path $env:TEMP ("{0}-envoyer-rapport.url" -f $ProvConfig.OrgName)
            }
            @"
[InternetShortcut]
URL=$mailto
"@ | Set-Content -LiteralPath $urlFile -Encoding ASCII -Force
            try { Start-Process explorer.exe -ArgumentList $urlFile } catch {}
            Write-Status -Name "Envoi email" -Status INFO -Detail "mailto non auto — ouvrir $urlFile" -NoCount
            Write-Host ("     Aucun client mail par defaut, ou bloque par UAC admin.") -ForegroundColor Yellow
            Write-Host ("     Double-clique : $urlFile") -ForegroundColor Yellow
            Write-Host ("     Ou cree un mail vers : $to") -ForegroundColor Yellow
        }
        Write-Host ("     Glissez les pieces jointes depuis le dossier ouvert :") -ForegroundColor Yellow
        foreach ($f in $attachments) { Write-Host ("       $f") -ForegroundColor DarkYellow }
        if ($script:ProvUsbReportDir) {
            Write-Host ("     Miroir cle USB : $($script:ProvUsbReportDir)") -ForegroundColor DarkCyan
        }
    } catch {
        Write-Status -Name "Envoi email" -Status FAIL -Detail $_.Exception.Message
        Write-Host ("     Rapport local : $($Export.Txt)") -ForegroundColor Yellow
        Write-Host ("     Destinataires  : $to") -ForegroundColor Yellow
    }
}

function Test-ProvPathOnRemovableDrive {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    try {
        $full = $Path
        if (Test-Path -LiteralPath $Path) {
            $full = (Resolve-Path -LiteralPath $Path).Path
        }
        $root = [System.IO.Path]::GetPathRoot($full)
        if (-not $root) { return $false }
        $letter = $root.TrimEnd('\')
        $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$letter'" -ErrorAction SilentlyContinue
        if (-not $disk) { return $false }
        # 2 = Removable ; parfois USB Fixed : on accepte aussi si volume porte scripts\offline
        if ($disk.DriveType -eq 2) { return $true }
        if (Test-ProvOfflineRootValid (Join-Path $letter "scripts\offline")) { return $true }
        if (Test-ProvOfflineRootValid (Join-Path $letter "offline")) { return $true }
        return $false
    } catch {
        return $false
    }
}

function Get-ProvUsbScriptsRoot {
    <# Racine scripts\ sur la cle (BitLocker unlock = inscriptible). #>
    if ($ScriptDir -and (Test-Path -LiteralPath $ScriptDir) -and (Test-ProvPathOnRemovableDrive -Path $ScriptDir)) {
        try { return (Resolve-Path -LiteralPath $ScriptDir).Path } catch { return $ScriptDir }
    }
    foreach ($off in @(Get-ProvOfflineDirsOnly)) {
        try {
            $parent = Split-Path -Parent $off
            if (-not $parent) { continue }
            if (-not (Test-ProvPathOnRemovableDrive -Path $parent)) { continue }
            if (-not (Test-ProvPathWritable -Dir $parent)) { continue }
            $leaf = Split-Path -Leaf $parent
            if ($leaf -eq 'scripts') { return (Resolve-Path -LiteralPath $parent).Path }
            $scripts = Join-Path $parent "scripts"
            if (Test-Path -LiteralPath $scripts) { return (Resolve-Path -LiteralPath $scripts).Path }
            return (Resolve-Path -LiteralPath $parent).Path
        } catch {}
    }
    return $null
}

function Sync-ProvReportsToUsb {
    <#
    Miroir Documents → cle\scripts\<UsbReportsFolder>\<PC>-<stamp>\
    Mailto conserve (SI distant). USB = archive terrain operateur.
    #>
    param($Export)
    if (-not $Export) { return $null }

    $stamp = if ($LogDate) { $LogDate } else { Get-Date -Format "yyyyMMdd-HHmmss" }
    $pc = if ($script:ProvComputerName) { $script:ProvComputerName } else { $env:COMPUTERNAME }
    $safe = ($pc -replace '[^\w\-]', '_')

    $usbRoot = Get-ProvUsbScriptsRoot
    if (-not $usbRoot) {
        Write-Status -Name "Rapports USB" -Status SKIP -Detail "cle non detectee"
        return $null
    }

    $dest = Join-Path $usbRoot (Join-Path $ProvConfig.UsbReportsFolder ("{0}-{1}" -f $safe, $stamp))
    if (-not (Test-ProvPathWritable -Dir $dest -Create)) {
        Write-Status -Name "Rapports USB" -Status SKIP -Detail "ecriture refusee (deverrouiller BitLocker ?)"
        return $null
    }

    $copied = 0
    foreach ($f in @(Get-ProvReportAttachments -Export $Export)) {
        try {
            Copy-Item -LiteralPath $f -Destination (Join-Path $dest (Split-Path $f -Leaf)) -Force -ErrorAction Stop
            $copied++
        } catch {
            Write-Log ("USB copy fail: {0} — {1}" -f $f, $_.Exception.Message) -Level WARN
        }
    }

    $note = @"
Rapports $($ProvConfig.OrgName) — miroir USB
=========================
PC     : $pc
Stamp  : $stamp
Source : Documents\$($ProvConfig.ReportFolderName) (poste)
Mailto : conserve (SI) — cette copie = archive terrain sur la cle BitLocker
"@
    try {
        Set-Content -Path (Join-Path $dest $ProvConfig.NoteFileName) -Value $note -Encoding UTF8 -Force
    } catch {}

    $script:ProvUsbReportDir = $dest
    if ($copied -gt 0) {
        Write-Status -Name "Rapports USB" -Status OK -Detail ("$copied fichier(s) → $dest")
    } else {
        Write-Status -Name "Rapports USB" -Status SKIP -Detail "aucun fichier a copier"
    }
    return $dest
}

function Invoke-ProvPostDeployCleanup {
    <#
    TOUTE FIN : apres email / resume.
    - ghost offline Bureau
    - confirmation EXPLICITE pour supprimer deploy-apps.* du poste (PAS la cle)
    - vider la corbeille
    Aussi dispo en secours : start.bat sur la cle (menu > Nettoyage).
    #>
    Write-Host ""
    Write-Host ("  " + $UI.Rule) -ForegroundColor DarkCyan
    Write-Host ("   NETTOYAGE POSTE  " + $UI.DotMid + "  derniere etape") -ForegroundColor Cyan
    Write-Host ("  " + $UI.Rule) -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host ("  " + $UI.VLine + "  Terminez d'abord l'envoi du mail (mailto),") -ForegroundColor Yellow
    Write-Host ("  " + $UI.VLine + "  puis revenez ICI pour le nettoyage.") -ForegroundColor Yellow
    Write-Host ""
    [void](Read-Host "     Entree quand le mail est pret / envoye — on continue le nettoyage")

    $desk = [Environment]::GetFolderPath("Desktop")

    # Ghost offline vide
    foreach ($ghost in @(
        (Join-Path $desk "offline"),
        (Join-Path $desk "scripts\offline")
    )) {
        if (-not (Test-Path -LiteralPath $ghost -PathType Container)) { continue }
        if (Test-ProvOfflineRootValid -Path $ghost) { continue }
        $kids = @(Get-ChildItem -LiteralPath $ghost -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notmatch '^\._' })
        if ($kids.Count -eq 0) {
            try {
                Remove-Item -LiteralPath $ghost -Force -Recurse -ErrorAction Stop
                Write-Status -Name "Ghost offline" -Status OK -Detail "supprime: $ghost" -NoCount
            } catch {}
        }
    }

    if (Test-ProvPathOnRemovableDrive -Path $ScriptDir) {
        Write-Status -Name "Cle USB" -Status INFO -Detail "lance depuis la cle — contenu USB intact" -NoCount
    }

    $targets = New-Object System.Collections.Generic.List[string]
    foreach ($name in @("deploy-apps.ps1", "deploy-apps.bat", "cleanup-deploy.ps1")) {
        foreach ($base in @($desk, (Join-Path $desk "scripts"), $ScriptDir)) {
            if (-not $base) { continue }
            $p = Join-Path $base $name
            if (-not (Test-Path -LiteralPath $p)) { continue }
            if (Test-ProvPathOnRemovableDrive -Path $p) { continue }  # jamais la cle
            if (-not $targets.Contains($p)) { [void]$targets.Add($p) }
        }
    }

    # Ne pas proposer cleanup-deploy.* s'il est le seul outil (on le laisse si lance depuis lui-meme sur bureau — OK)
    if ($targets.Count -eq 0) {
        Write-Status -Name "Nettoyage script" -Status SKIP -Detail "aucun deploy-apps sur le Bureau (deja clean ?)"
        Write-Host ("  " + $UI.Arrow + " Secours : start.bat sur la cle (menu > Nettoyage) si besoin.") -ForegroundColor DarkGray
        return
    }

    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " Fichiers a supprimer du POSTE " + ($UI.HLine * 20)) -ForegroundColor Magenta
    foreach ($t in $targets) {
        Write-Host ("  " + $UI.VLine + "  $t") -ForegroundColor Gray
    }
    Write-Host ("  " + $UI.VLine + "  Conserves : Documents\$($ProvConfig.ReportFolderName) + TOUTE la cle USB") -ForegroundColor DarkGray
    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
    Write-Host ""
    Write-Host ("     Tapez  OUI  pour supprimer + vider la corbeille") -ForegroundColor Yellow
    Write-Host ("     Tapez  non  (ou Entree) pour garder les fichiers sur le poste") -ForegroundColor DarkGray
    $ans = (Read-Host "     Confirmation").Trim().ToLowerInvariant()
    if ($ans -notin @('o', 'oui', 'y', 'yes')) {
        Write-Status -Name "Nettoyage script" -Status SKIP -Detail "non confirme — fichiers conserves"
        Write-Host ("  " + $UI.Arrow + " Plus tard : start.bat sur la cle USB (menu > Nettoyage)") -ForegroundColor Cyan
        return
    }

    foreach ($t in $targets) {
        try {
            Remove-Item -LiteralPath $t -Force -ErrorAction Stop
            Write-Status -Name "Supprime" -Status OK -Detail $t -NoCount
        } catch {
            Write-Status -Name "Supprime" -Status FAIL -Detail "$t — $($_.Exception.Message)" -NoCount
        }
    }

    $deskScripts = Join-Path $desk "scripts"
    if ((Test-Path $deskScripts) -and -not (Test-ProvPathOnRemovableDrive $deskScripts)) {
        $left = @(Get-ChildItem -LiteralPath $deskScripts -Force -EA SilentlyContinue | Where-Object { $_.Name -notmatch '^\._' })
        if ($left.Count -eq 0) {
            Remove-Item -LiteralPath $deskScripts -Force -Recurse -EA SilentlyContinue
        }
    }

    try {
        Clear-RecycleBin -Force -ErrorAction Stop
        Write-Status -Name "Corbeille" -Status OK -Detail "videe"
    } catch {
        try {
            $shell = New-Object -ComObject Shell.Application
            $shell.NameSpace(0xA).Items() | ForEach-Object { Remove-Item $_.Path -Recurse -Force -EA SilentlyContinue }
            Write-Status -Name "Corbeille" -Status OK -Detail "videe (shell)"
        } catch {
            Write-Status -Name "Corbeille" -Status FAIL -Detail $_.Exception.Message
        }
    }

    if ($script:ProvReportDir) {
        Write-Host ("  " + $UI.Arrow + " Rapports conserves : $($script:ProvReportDir)") -ForegroundColor Cyan
    }
}


function Find-ProvMigrationPacks {
    <# Cherche migrationWin11_* sur cles USB / lecteurs (<FilePrefix>-migration.json ou Desktop|Documents|Downloads). #>
    $packs = New-Object System.Collections.Generic.List[object]
    $roots = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($d in @(Get-CimInstance Win32_LogicalDisk -EA SilentlyContinue)) {
            if (-not $d.DeviceID) { continue }
            if ($d.DriveType -in @(2, 3) -and $d.DeviceID -ne $env:SystemDrive) {
                [void]$roots.Add(($d.DeviceID.TrimEnd('\') + '\'))
            }
        }
    } catch {}
    # Aussi a cote du script / offline parent
    foreach ($extra in @($ScriptDir, (Join-Path $ScriptDir ".."))) {
        try {
            if ($extra -and (Test-Path $extra)) { [void]$roots.Add((Resolve-Path $extra).Path) }
        } catch {}
    }

    foreach ($root in ($roots | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        $dirs = @(Get-ChildItem -LiteralPath $root -Directory -EA SilentlyContinue |
            Where-Object { $_.Name -match '^migrationWin11_' })
        # Un niveau plus bas (ex. E:\MIGRATE\migrationWin11_...)
        if ($dirs.Count -eq 0) {
            Get-ChildItem -LiteralPath $root -Directory -EA SilentlyContinue | ForEach-Object {
                $sub = Get-ChildItem -LiteralPath $_.FullName -Directory -EA SilentlyContinue |
                    Where-Object { $_.Name -match '^migrationWin11_' }
                foreach ($s in $sub) { $dirs += $s }
            }
        }
        foreach ($dir in $dirs) {
            $manifest = Join-Path $dir.FullName ("{0}-migration.json" -f $ProvConfig.FilePrefix)
            $pwdCsvs = @(Get-ChildItem -LiteralPath $dir.FullName -Filter "chrome-passwords*.csv" -File -EA SilentlyContinue)
            $bmHtmls = @(Get-ChildItem -LiteralPath $dir.FullName -Filter "chrome-bookmarks*.html" -File -EA SilentlyContinue)
            $bmJsons = @(Get-ChildItem -LiteralPath $dir.FullName -Filter "chrome-bookmarks*.json" -File -EA SilentlyContinue)
            $hasContent = (Test-Path (Join-Path $dir.FullName "Desktop")) -or
                          (Test-Path (Join-Path $dir.FullName "Documents")) -or
                          (Test-Path (Join-Path $dir.FullName "Downloads")) -or
                          ($pwdCsvs.Count -gt 0) -or ($bmHtmls.Count -gt 0) -or ($bmJsons.Count -gt 0)
            if (-not $hasContent -and -not (Test-Path $manifest)) { continue }
            $label = $dir.Name
            $prenom = ""; $nom = ""
            if (Test-Path $manifest) {
                try {
                    $m = Get-Content -LiteralPath $manifest -Raw -Encoding UTF8 | ConvertFrom-Json
                    if ($m.prenom) { $prenom = [string]$m.prenom }
                    if ($m.nom) { $nom = [string]$m.nom }
                    if ($prenom -or $nom) { $label = "$($dir.Name)  ($prenom $nom)".Trim() }
                } catch {}
            }
            [void]$packs.Add([pscustomobject]@{
                Path    = $dir.FullName
                Name    = $dir.Name
                Label   = $label
                HasCsv  = [bool]($pwdCsvs.Count -gt 0)
            })
        }
    }
    return @($packs | Sort-Object Path -Unique)
}

function Import-ProvMigrationUserFiles {
    <# Debut de deploy : importer Bureau/Documents/Telechargements depuis migrationWin11_*. #>
    Write-Section -Title "Import migration (ancien poste)" -Step "1b/8"
    $packs = @(Find-ProvMigrationPacks)
    if ($packs.Count -eq 0) {
        Write-Status -Name "Migration USB" -Status INFO -Detail "aucun dossier migrationWin11_* — skip" -NoCount
        $manual = (Read-Host "     Chemin migrationWin11_... [Entree = passer]").Trim().Trim('"')
        if (-not $manual) { return }
        if (-not (Test-Path -LiteralPath $manual -PathType Container)) {
            Write-Status -Name "Migration" -Status FAIL -Detail "chemin invalide"
            return
        }
        $packs = @([pscustomobject]@{
            Path = $manual; Name = (Split-Path $manual -Leaf); Label = (Split-Path $manual -Leaf)
            HasCsv = [bool]@(Get-ChildItem -LiteralPath $manual -Filter "chrome-passwords*.csv" -File -EA SilentlyContinue).Count
        })
    }

    Write-Host ""
    Write-Host ("  " + $UI.VLine + "  Packs detectes :") -ForegroundColor Gray
    $i = 0
    foreach ($p in $packs) {
        $i++
        $csv = if ($p.HasCsv) { " + mdp Chrome" } else { "" }
        Write-Host ("  " + $UI.VLine + "  [$i] $($p.Label)$csv") -ForegroundColor Cyan
        Write-Host ("  " + $UI.VLine + "      $($p.Path)") -ForegroundColor DarkGray
    }
    Write-Host ""

    if (-not (Read-ProvYesNo -Prompt "Importer les fichiers utilisateur maintenant ?" -Default Yes)) {
        Write-Status -Name "Migration fichiers" -Status SKIP -Detail "refuse par l'operateur"
        # Garder quand meme le chemin si un seul pack (pour mdp Chrome plus tard)
        if ($packs.Count -eq 1) { $script:ProvMigrationRoot = $packs[0].Path }
        return
    }

    $chosen = $null
    if ($packs.Count -eq 1) {
        $chosen = $packs[0]
    } else {
        $pick = (Read-Host "     Numero du pack [1]").Trim()
        if ([string]::IsNullOrWhiteSpace($pick)) { $pick = "1" }
        $n = 0
        [void][int]::TryParse($pick, [ref]$n)
        if ($n -lt 1 -or $n -gt $packs.Count) { $n = 1 }
        $chosen = $packs[$n - 1]
    }

    $script:ProvMigrationRoot = $chosen.Path
    Write-Status -Name "Pack migration" -Status OK -Detail $chosen.Path

    $map = @(
        @{ Src = "Desktop";   Dst = [Environment]::GetFolderPath("Desktop");     Label = "Bureau" }
        @{ Src = "Documents"; Dst = [Environment]::GetFolderPath("MyDocuments"); Label = "Documents" }
        @{ Src = "Downloads"; Dst = (Join-Path $env:USERPROFILE "Downloads");    Label = "Telechargements" }
    )

    foreach ($m in $map) {
        $src = Join-Path $chosen.Path $m.Src
        if (-not (Test-Path -LiteralPath $src)) {
            Write-Status -Name $m.Label -Status SKIP -Detail "absent dans le pack" -NoCount
            continue
        }
        if (-not (Test-Path -LiteralPath $m.Dst)) {
            New-Item -ItemType Directory -Path $m.Dst -Force | Out-Null
        }
        Write-Status -Name $m.Label -Status RUN -Detail "copie $src -> $($m.Dst)" -NoCount
        $p = Start-Process robocopy.exe -ArgumentList @(
            $src, $m.Dst, "/E", "/COPY:DAT", "/R:1", "/W:1", "/NFL", "/NDL", "/NJH", "/NJS", "/NP"
        ) -Wait -PassThru -WindowStyle Hidden
        if ($p.ExitCode -le 7) {
            Write-Status -Name $m.Label -Status OK -Detail $m.Dst
        } else {
            Write-Status -Name $m.Label -Status FAIL -Detail "robocopy exit $($p.ExitCode)"
        }
    }

    Write-Status -Name "Migration fichiers" -Status OK -Detail "import termine (mdp Chrome apres install Chrome)"
}

function ConvertTo-ProvChromeBookmarksHtml {
    <#
      JSON Bookmarks Chrome → HTML Netscape.
      bookmark_bar avec PERSONAL_TOOLBAR_FOLDER=true + enfants directs
      (sinon Chrome cree un dossier "Barre de favoris").
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

function Resolve-ProvChromeProfileDirectoryFromFile {
    <#
      chrome-passwords.csv / chrome-bookmarks.html → Default
      chrome-passwords-Profile_1.csv → "Profile 1"
    #>
    param([string]$FileName)
    $base = [IO.Path]::GetFileNameWithoutExtension($FileName)
    if ($base -match '(?i)^chrome-(?:passwords|bookmarks)$') { return 'Default' }
    if ($base -match '(?i)^chrome-(?:passwords|bookmarks)-(.+)$') {
        $suf = $Matches[1]
        if ($suf -match '(?i)^Profile_(\d+)$') { return ("Profile {0}" -f $Matches[1]) }
        if ($suf -match '(?i)^Profile\s*(\d+)$') { return ("Profile {0}" -f $Matches[1]) }
        return ($suf -replace '_', ' ')
    }
    return 'Default'
}

function Start-ProvChromeForImport {
    param(
        [string]$ChromeExe,
        [string]$ProfileDirectory,
        [string]$Url
    )
    if (-not $ChromeExe) { return $false }
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ChromeExe
        $psi.Arguments = ('--profile-directory="{0}" {1}' -f $ProfileDirectory, $Url)
        $psi.UseShellExecute = $true
        [void][System.Diagnostics.Process]::Start($psi)
        return $true
    } catch {
        try {
            Start-Process -FilePath $ChromeExe -ArgumentList @("--profile-directory=$ProfileDirectory", $Url)
            return $true
        } catch { return $false }
    }
}

function Import-ProvChromePasswordsFromMigration {
    <# Apres install Chrome : import CSV mdp + favoris HTML, bloque jusqu'a OK. Multi-profil. #>
    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " Import Chrome (mdp + favoris) " + ($UI.HLine * 20)) -ForegroundColor Magenta

    # Retrouver le pack si besoin
    if (-not $script:ProvMigrationRoot -or -not (Test-Path -LiteralPath $script:ProvMigrationRoot)) {
        $packs = @(Find-ProvMigrationPacks)
        $withChrome = @($packs | Where-Object {
            $_.HasCsv -or
            (Get-ChildItem -LiteralPath $_.Path -Filter "chrome-bookmarks*.html" -File -EA SilentlyContinue | Select-Object -First 1) -or
            (Get-ChildItem -LiteralPath $_.Path -Filter "chrome-bookmarks*.json" -File -EA SilentlyContinue | Select-Object -First 1)
        })
        if ($withChrome.Count -eq 1) {
            $script:ProvMigrationRoot = $withChrome[0].Path
        } elseif ($withChrome.Count -gt 1) {
            Write-Host ("  " + $UI.VLine + "  Plusieurs packs Chrome :") -ForegroundColor Gray
            $i = 0
            foreach ($p in $withChrome) {
                $i++
                Write-Host ("  " + $UI.VLine + "  [$i] $($p.Label)") -ForegroundColor Cyan
            }
            $pick = (Read-Host "     Pack [1 / S = skip tout]").Trim().ToLowerInvariant()
            if ($pick -in @('s', 'skip')) {
                Write-Status -Name "Import Chrome" -Status SKIP -Detail "ignore"
                Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
                return
            }
            if ([string]::IsNullOrWhiteSpace($pick)) { $pick = "1" }
            $n = 0; [void][int]::TryParse($pick, [ref]$n)
            if ($n -ge 1 -and $n -le $withChrome.Count) {
                $script:ProvMigrationRoot = $withChrome[$n - 1].Path
            } else {
                Write-Host ("  " + $UI.VLine + "  Choix invalide — skip import Chrome") -ForegroundColor Yellow
                Write-Status -Name "Import Chrome" -Status SKIP -Detail "choix pack invalide"
                Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
                return
            }
        }
    }

    $root = $script:ProvMigrationRoot
    $csvList = @()
    $htmlList = @()
    if ($root -and (Test-Path -LiteralPath $root)) {
        $csvList = @(Get-ChildItem -LiteralPath $root -Filter "chrome-passwords*.csv" -File -EA SilentlyContinue |
            Sort-Object { if ($_.Name -eq 'chrome-passwords.csv') { 0 } else { 1 } }, Name)
        # Regenerer HTML depuis JSON (corrige packs anciens : dossier "Barre de favoris")
        $jsonBms = @(Get-ChildItem -LiteralPath $root -Filter "chrome-bookmarks*.json" -File -EA SilentlyContinue)
        foreach ($jf in $jsonBms) {
            $htmlTarget = [IO.Path]::ChangeExtension($jf.FullName, '.html')
            if (ConvertTo-ProvChromeBookmarksHtml -BookmarksJsonPath $jf.FullName -OutHtmlPath $htmlTarget) {
                Write-Host ("  " + $UI.VLine + "  Favoris HTML regenere : $($jf.Name) → $([IO.Path]::GetFileName($htmlTarget))") -ForegroundColor DarkGray
            }
        }
        $htmlList = @(Get-ChildItem -LiteralPath $root -Filter "chrome-bookmarks*.html" -File -EA SilentlyContinue |
            Sort-Object { if ($_.Name -eq 'chrome-bookmarks.html') { 0 } else { 1 } }, Name)
        if ($htmlList.Count -eq 0 -and $jsonBms.Count -gt 0) {
            Write-Host ("  " + $UI.VLine + "  JSON favoris presents mais conversion HTML echouee") -ForegroundColor Yellow
        }
    }

    if ($csvList.Count -eq 0 -and $htmlList.Count -eq 0) {
        Write-Host ("  " + $UI.VLine + "  Pas de CSV / favoris migration trouves.") -ForegroundColor Yellow
        if (-not (Read-ProvYesNo -Prompt "Continuer sans import Chrome ?" -Default Yes)) {
            $manual = (Read-Host "     Chemin dossier migrationWin11_...").Trim().Trim('"')
            if ($manual -and (Test-Path $manual)) {
                $script:ProvMigrationRoot = $manual
                Import-ProvChromePasswordsFromMigration
                return
            }
        }
        Write-Status -Name "Import Chrome" -Status SKIP -Detail "rien a importer"
        Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
        return
    }

    if ($csvList.Count + $htmlList.Count -gt 2) {
        Write-Host ("  " + $UI.VLine + "  Multi-profil : $($csvList.Count) CSV mdp, $($htmlList.Count) HTML favoris") -ForegroundColor Cyan
        Write-Host ("  " + $UI.VLine + "  Chaque fichier ouvre le profil Chrome correspondant (Default / Profile N).") -ForegroundColor DarkGray
        Write-Host ("  " + $UI.VLine + "  Sur PC neuf : creer le profil avant si besoin, ou tout importer dans Default.") -ForegroundColor DarkGray
    }

    $desk = [Environment]::GetFolderPath("Desktop")
    $chrome = Find-ChromeExe
    if (-not $chrome) {
        Write-Host ("  " + $UI.VLine + "  Chrome introuvable — installez-le d'abord (module chrome-install).") -ForegroundColor Yellow
    }

    # --- Mots de passe (1+ profils) ---
    if ($csvList.Count -gt 0) {
        $pi = 0
        foreach ($csv in $csvList) {
            $pi++
            $profileDir = Resolve-ProvChromeProfileDirectoryFromFile -FileName $csv.Name
            $label = if ($csv.Name -eq 'chrome-passwords.csv') { 'Default' } else { $csv.BaseName -replace '^chrome-passwords-?', '' }
            $safeTag = if ($profileDir -eq 'Default') { 'Default' } else { ($profileDir -replace '\s', '_') }
            $deskCsv = Join-Path $desk ("{0}-chrome-passwords-{1}.csv" -f $ProvConfig.FilePrefix, $safeTag)
            try { Copy-Item -LiteralPath $csv.FullName -Destination $deskCsv -Force } catch { $deskCsv = $csv.FullName }
            Write-Host ""
            Write-Host ("  " + $UI.VLine + "  MDP [$pi/$($csvList.Count)] $label  →  profil Chrome « $profileDir »") -ForegroundColor Cyan
            Write-Host ("  " + $UI.VLine + "  → $deskCsv") -ForegroundColor DarkCyan
            Write-Host ("  " + $UI.VLine + "  Chrome → Parametres → Mots de passe → Importer") -ForegroundColor Yellow
            [void](Start-ProvChromeForImport -ChromeExe $chrome -ProfileDirectory $profileDir -Url "chrome://password-manager/settings")
            try { Start-Process explorer.exe -ArgumentList "/select,`"$deskCsv`"" } catch {}
            $okPwd = $false
            while (-not $okPwd) {
                $okPwd = Read-ProvYesNo -Prompt "Mdp [$label / $profileDir] importes — OK ?" -Default Yes
                if (-not $okPwd) {
                    Write-Host ("     Terminez l'import mdp puis O.") -ForegroundColor Yellow
                    [void](Start-ProvChromeForImport -ChromeExe $chrome -ProfileDirectory $profileDir -Url "chrome://password-manager/settings")
                }
            }
            Write-Status -Name ("Import mdp ($label)") -Status OK -Detail "profil=$profileDir"
            if ($deskCsv -ne $csv.FullName -and (Test-Path $deskCsv)) {
                if (Read-ProvYesNo -Prompt "Supprimer copie Bureau ($([IO.Path]::GetFileName($deskCsv))) ?" -Default Yes) {
                    Remove-Item $deskCsv -Force -EA SilentlyContinue
                }
            }
        }
    } else {
        Write-Status -Name "Import mdp Chrome" -Status SKIP -Detail "pas de chrome-passwords*.csv" -NoCount
    }

    # --- Favoris (1+ profils) ---
    if ($htmlList.Count -gt 0) {
        $bi = 0
        foreach ($htmlBm in $htmlList) {
            $bi++
            $profileDir = Resolve-ProvChromeProfileDirectoryFromFile -FileName $htmlBm.Name
            $label = if ($htmlBm.Name -eq 'chrome-bookmarks.html') { 'Default' } else { $htmlBm.BaseName -replace '^chrome-bookmarks-?', '' }
            $safeTag = if ($profileDir -eq 'Default') { 'Default' } else { ($profileDir -replace '\s', '_') }
            $deskHtml = Join-Path $desk ("{0}-chrome-bookmarks-{1}.html" -f $ProvConfig.FilePrefix, $safeTag)
            try { Copy-Item -LiteralPath $htmlBm.FullName -Destination $deskHtml -Force } catch { $deskHtml = $htmlBm.FullName }
            Write-Host ""
            Write-Host ("  " + $UI.VLine + "  FAVORIS [$bi/$($htmlList.Count)] $label  →  profil Chrome « $profileDir »") -ForegroundColor Cyan
            Write-Host ("  " + $UI.VLine + "  → $deskHtml") -ForegroundColor DarkCyan
            Write-Host ("  " + $UI.VLine + "  Chrome → chrome://bookmarks → ⋮ → Importer") -ForegroundColor Yellow
            [void](Start-ProvChromeForImport -ChromeExe $chrome -ProfileDirectory $profileDir -Url "chrome://bookmarks/")
            try { Start-Process explorer.exe -ArgumentList "/select,`"$deskHtml`"" } catch {}
            $okBm = $false
            while (-not $okBm) {
                $okBm = Read-ProvYesNo -Prompt "Favoris [$label / $profileDir] importes — OK ?" -Default Yes
                if (-not $okBm) {
                    Write-Host ("     Terminez l'import favoris puis O.") -ForegroundColor Yellow
                    [void](Start-ProvChromeForImport -ChromeExe $chrome -ProfileDirectory $profileDir -Url "chrome://bookmarks/")
                }
            }
            Write-Status -Name ("Import favoris ($label)") -Status OK -Detail "profil=$profileDir"
            if ($deskHtml -ne $htmlBm.FullName -and (Test-Path $deskHtml)) {
                if (Read-ProvYesNo -Prompt "Supprimer copie Bureau HTML ?" -Default Yes) {
                    Remove-Item $deskHtml -Force -EA SilentlyContinue
                }
            }
        }
    } else {
        Write-Status -Name "Import favoris Chrome" -Status SKIP -Detail "pas de chrome-bookmarks*.html" -NoCount
    }

    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
}


# ============================================================================
#  MODE MODULES (retouches sans full deploy)
# ============================================================================

function Get-ProvDeployModuleCatalog {
    # Ordre catalogue = ordre d'affichage ; l'execution est reordonnee (Sort-ProvModuleIdsByDeps).
    @(
        [pscustomobject]@{ Id = 'migration-files'; Group = 'Chrome / migration'; Label = 'Import fichiers migration (USB)'; NeedsWinget = $false; Priority = 10; Run = { Import-ProvMigrationUserFiles } }
        [pscustomobject]@{ Id = 'chrome-install';  Group = 'Chrome / migration'; Label = 'Installer / reparer Google Chrome'; NeedsWinget = $false; Priority = 20; Run = { Install-ChromeWithFallback | Out-Null } }
        [pscustomobject]@{ Id = 'chrome-import';   Group = 'Chrome / migration'; Label = 'Import mdp + favoris Chrome'; NeedsWinget = $false; Priority = 30; Run = { Import-ProvChromePasswordsFromMigration } }
        [pscustomobject]@{ Id = 'chrome-default';  Group = 'Personnalisation'; Label = 'Chrome navigateur + PDF par defaut'; NeedsWinget = $false; Priority = 40; Run = { Set-ChromeAsDefault } }
        [pscustomobject]@{ Id = 'startup';         Group = 'Personnalisation'; Label = 'Apps au demarrage (Chrome ON, Firefox OFF)'; NeedsWinget = $false; Priority = 50; Run = { Set-ProvStartupApps } }
        [pscustomobject]@{ Id = 'tweaks';          Group = 'Personnalisation'; Label = 'Optimisations Win11 (perfs / tips)'; NeedsWinget = $false; Priority = 51; Run = { Set-ProvPerformanceTweaks } }
        [pscustomobject]@{ Id = 'widgets';         Group = 'Personnalisation'; Label = 'Meteo / Widgets barre des taches OFF'; NeedsWinget = $false; Priority = 52; Run = { Disable-ProvTaskbarWidgets } }
        [pscustomobject]@{ Id = 'taskbar';         Group = 'Personnalisation'; Label = ("Barre des taches {0} (pins)" -f $ProvConfig.OrgName); NeedsWinget = $false; Priority = 53; Run = { Set-ProvTaskbar } }
        [pscustomobject]@{ Id = 'desktop';         Group = 'Personnalisation'; Label = 'Rangement Bureau (_Inbox)'; NeedsWinget = $false; Priority = 54; Run = { Set-ProvDesktopOrganize } }
        [pscustomobject]@{ Id = 'bloat';           Group = 'Personnalisation'; Label = 'Retirer bloatware'; NeedsWinget = $false; Priority = 55; Run = { Remove-ProvBloatware } }
        [pscustomobject]@{ Id = 'network';         Group = 'Personnalisation'; Label = 'Lecteurs reseau + imprimantes'; NeedsWinget = $false; Priority = 56; Run = { Set-ProvNetworkResources } }
        [pscustomobject]@{ Id = 'anydesk-pwd';     Group = 'Personnalisation'; Label = 'Mot de passe AnyDesk unattended'; NeedsWinget = $false; Priority = 57; Run = { Set-ProvAnyDeskPassword } }
        [pscustomobject]@{ Id = 'apps';            Group = 'Installations'; Label = 'Apps winget/direct (Firefox, Discord…)'; NeedsWinget = $false; Priority = 60; Run = { foreach ($a in $script:AppsWinget) { Install-WingetApp -Name $a.Name -Id $a.Id | Out-Null } } }
        [pscustomobject]@{ Id = 'anydesk';         Group = 'Installations'; Label = 'Installer AnyDesk (offline/direct)'; NeedsWinget = $false; Priority = 61; Run = { Install-AnyDeskDirect | Out-Null } }
        [pscustomobject]@{ Id = 'java';            Group = 'Installations'; Label = 'Java (Oracle/Temurin + fallback)'; NeedsWinget = $false; Priority = 62; Run = { Install-JavaRuntime } }
        [pscustomobject]@{ Id = 'metier';          Group = 'Installations'; Label = 'Apps metier offline (voir config)'; NeedsWinget = $false; Priority = 70; Run = { Install-ProvLobApps } }
        [pscustomobject]@{ Id = 'office';          Group = 'Installations'; Label = 'Microsoft 365 Apps (winget → ODT)'; NeedsWinget = $false; Priority = 80; Run = { $script:ProvModuleOfficeOk = [bool](Install-Microsoft365) } }
        [pscustomobject]@{ Id = 'office-login';    Group = 'Installations'; Label = 'Connexion / login Microsoft 365'; NeedsWinget = $false; Priority = 90; Run = {
            if (Get-Command Start-OfficeSignIn -EA SilentlyContinue) {
                Start-OfficeSignIn
            }
        } }
        [pscustomobject]@{ Id = 'identity';        Group = 'Identite / rapports'; Label = 'Identite PC + user local'; NeedsWinget = $false; Priority = 5; Run = { Set-ProvComputerIdentity; New-ProvLocalUser } }
        [pscustomobject]@{ Id = 'update';          Group = 'Identite / rapports'; Label = 'Windows Update'; NeedsWinget = $false; Priority = 100; Run = { Start-ProvWindowsUpdate } }
        [pscustomobject]@{ Id = 'intune';          Group = 'Identite / rapports'; Label = 'Check Entra ID / Intune'; NeedsWinget = $false; Priority = 101; Run = { Test-ProvIntuneCompliance | Out-Null } }
        [pscustomobject]@{ Id = 'report';          Group = 'Identite / rapports'; Label = 'Inventaire + rapport + mailto'; NeedsWinget = $false; Priority = 110; Run = { $e = Export-ProvInventory; Sync-ProvReportsToUsb -Export $e | Out-Null; Send-ProvReport -Export $e } }
    )
}

function Expand-ProvModuleIds {
    param([string[]]$Raw)
    $aliases = @{
        perso   = @('chrome-default', 'widgets', 'tweaks', 'startup', 'taskbar')
        ui      = @('chrome-default', 'widgets', 'tweaks', 'startup', 'taskbar')
        windows = @('chrome-default', 'widgets', 'tweaks', 'startup', 'taskbar', 'desktop', 'bloat')
        chrome  = @('chrome-install', 'chrome-import', 'chrome-default')
    }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($r in @($Raw)) {
        if ([string]::IsNullOrWhiteSpace($r)) { continue }
        foreach ($part in ($r -split '[,;\s]+' | Where-Object { $_ })) {
            $key = $part.Trim().ToLowerInvariant()
            if ($aliases.ContainsKey($key)) {
                foreach ($a in $aliases[$key]) { if ($out -notcontains $a) { [void]$out.Add($a) } }
            } else {
                if ($out -notcontains $key) { [void]$out.Add($key) }
            }
        }
    }
    return @($out)
}

function Sort-ProvModuleIdsByDeps {
    <# Reordonne : identity → migration → chrome install → import → … #>
    param([string[]]$Ids, $Catalog)
    $prio = @{}
    foreach ($m in @($Catalog)) { $prio[$m.Id] = [int]$m.Priority }
    return @($Ids | Sort-Object {
        if ($prio.ContainsKey($_)) { $prio[$_] } else { 999 }
    }, { $_ })
}

function Show-ProvModulePicker {
    <# Multi-select : Haut/Bas, Espace, Entree, P=perso, A=tout, Echap=annuler. #>
    $catalog = @(Get-ProvDeployModuleCatalog)
    $sel = @{}
    foreach ($m in $catalog) { $sel[$m.Id] = $false }
    $i = 0
    $count = $catalog.Count

    function Show-Picker {
        Clear-Host
        Write-Host ""
        Write-Host ("  " + $UI.Rule) -ForegroundColor DarkCyan
        Write-Host ("   " + $ProvConfig.OrgName + "  " + $UI.DotMid + "  Modules / retouches") -ForegroundColor Cyan
        Write-Host ("   Sans full deploy — coche ce dont tu as besoin") -ForegroundColor DarkGray
        Write-Host ("  " + $UI.Rule) -ForegroundColor DarkCyan
        Write-Host ""
        Write-Host "  Haut/Bas  ·  Espace = cocher  ·  Entree = lancer" -ForegroundColor DarkGray
        Write-Host "  P = preset perso  ·  A = tout  ·  N = rien  ·  Echap = annuler" -ForegroundColor DarkGray
        Write-Host ""
        $lastGroup = ""
        for ($n = 0; $n -lt $count; $n++) {
            $m = $catalog[$n]
            if ($m.Group -ne $lastGroup) {
                Write-Host ("  --- " + $m.Group + " ---") -ForegroundColor Magenta
                $lastGroup = $m.Group
            }
            $mark = if ($sel[$m.Id]) { '[x]' } else { '[ ]' }
            $line = "  $mark  $($m.Label)  ($($m.Id))"
            if ($n -eq $i) {
                Write-Host $line -ForegroundColor Black -BackgroundColor Cyan
            } else {
                $col = if ($sel[$m.Id]) { 'Green' } else { 'Gray' }
                Write-Host $line -ForegroundColor $col
            }
        }
        $nSel = @($sel.Keys | Where-Object { $sel[$_] }).Count
        Write-Host ""
        Write-Host ("  Selection : $nSel module(s)") -ForegroundColor Cyan
        Write-Host ""
    }

    Show-Picker
    while ($true) {
        $k = [Console]::ReadKey($true)
        switch ($k.Key) {
            'UpArrow'   { $i = ($i - 1 + $count) % $count; Show-Picker }
            'DownArrow' { $i = ($i + 1) % $count; Show-Picker }
            'Spacebar'  {
                $id = $catalog[$i].Id
                $sel[$id] = -not $sel[$id]
                Show-Picker
            }
            'A' {
                foreach ($m in $catalog) { $sel[$m.Id] = $true }
                Show-Picker
            }
            'N' {
                foreach ($m in $catalog) { $sel[$m.Id] = $false }
                Show-Picker
            }
            'P' {
                foreach ($m in $catalog) { $sel[$m.Id] = $false }
                foreach ($id in (Expand-ProvModuleIds -Raw @('perso'))) { $sel[$id] = $true }
                Show-Picker
            }
            'Escape' { return @() }
            'Enter' {
                $chosen = @($catalog | Where-Object { $sel[$_.Id] } | ForEach-Object { $_.Id })
                if ($chosen.Count -eq 0) {
                    Write-Host "  Aucun module coche — Espace pour selectionner." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                    Show-Picker
                } else {
                    return $chosen
                }
            }
        }
    }
}

function Invoke-ProvModuleMode {
    param([string[]]$ModuleIds)

    $catalog = @(Get-ProvDeployModuleCatalog)
    $ids = Sort-ProvModuleIdsByDeps -Ids (Expand-ProvModuleIds -Raw $ModuleIds) -Catalog $catalog
    $byId = @{}
    foreach ($m in $catalog) { $byId[$m.Id] = $m }

    $unknown = @($ids | Where-Object { -not $byId.ContainsKey($_) })
    if ($unknown.Count -gt 0) {
        Write-Host ("  Modules inconnus : " + ($unknown -join ', ')) -ForegroundColor Red
        Write-Host "  Ids valides :" -ForegroundColor Yellow
        $catalog | ForEach-Object { Write-Host ("    $($_.Id) — $($_.Label)") -ForegroundColor DarkGray }
        return 1
    }

    Clear-Host
    Write-Host ""
    Write-Host ("  " + $UI.Rule) -ForegroundColor DarkCyan
    Write-Host ("   " + $ProvConfig.OrgName + "  " + $UI.DotMid + "  Mode modules") -ForegroundColor Cyan
    Write-Host ("   " + ($ids -join ' · ')) -ForegroundColor DarkGray
    Write-Host ("  " + $UI.Rule) -ForegroundColor DarkCyan
    Write-Host ""

    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Status -Name "Droits admin" -Status FAIL -Detail "relancer via start.bat"
        return 1
    }
    Write-Status -Name "Droits admin" -Status OK -NoCount
    Set-ProvConsoleMaximized

    # Winget : best-effort (Chrome/Office/apps ont des fallbacks offline/direct)
    if (Test-WingetAvailable) {
        Write-Status -Name "winget" -Status OK -Detail "$(winget --version 2>$null)" -NoCount
        Initialize-ProvWingetNetwork
    } else {
        Write-Status -Name "winget" -Status SKIP -Detail "absent — fallbacks offline/direct si dispo" -NoCount
    }

    $step = 0
    $total = $ids.Count
    foreach ($id in $ids) {
        $step++
        $m = $byId[$id]
        Write-Section -Title $m.Label -Step ("$step/$total")
        try {
            & $m.Run
        } catch {
            Write-Status -Name $m.Id -Status FAIL -Detail $_.Exception.Message
            Write-Log "Module $id FAIL: $($_.Exception.Message)" -Level ERR
        }
    }

    Write-Summary
    Write-Log "Modules fin. ids=$($ids -join ',') OK=$($script:Stats.OK) FAIL=$($script:Stats.Fail)" -Level INFO
    Write-Host ("  " + $UI.Arrow + " Mode modules termine (pas de cleanup / pas de reset etat deploy).") -ForegroundColor DarkGray
    Write-Host ""
    return 0
}


# ============================================================================
#  MAIN
# ============================================================================

try {
# --- Mode modules / retouches (pas de full deploy) ---
if ($ListModules) {
    Write-Host "Modules deploy-apps.ps1 :" -ForegroundColor Cyan
    Write-Host "  Alias : perso | ui | windows | chrome" -ForegroundColor DarkGray
    Get-ProvDeployModuleCatalog | ForEach-Object {
        Write-Host ("  [{0}] {1,-16} {2}" -f $_.Group, $_.Id, $_.Label)
    }
    return
}

$moduleIds = @()
if ($ModulePicker) {
    $moduleIds = @(Show-ProvModulePicker)
    if ($moduleIds.Count -eq 0) {
        Write-Host "  Annule." -ForegroundColor Yellow
        return
    }
} elseif ($Modules -and $Modules.Count -gt 0) {
    $moduleIds = @(Expand-ProvModuleIds -Raw $Modules)
}

if ($moduleIds.Count -gt 0) {
    if ($script:LogEnabled -and $LogFile) {
        try {
            $hdr = "$($ProvConfig.OrgName) modules $(Get-Date -Format o) : $($moduleIds -join ',')"
            Set-Content -Path $LogFile -Value $hdr -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
    }
    $null = Invoke-ProvModuleMode -ModuleIds $moduleIds
    return
}

Write-Banner

$logHeader = @"
============================================================
$($ProvConfig.OrgName) - Journal d'installation
Date        : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Machine     : $env:COMPUTERNAME
Utilisateur : $env:USERNAME
OS          : $([Environment]::OSVersion.VersionString)
ScriptDir   : $ScriptDir
ReportDir   : $($script:ProvReportDir)
LogFile     : $LogFile
============================================================
"@

if ($script:LogEnabled -and $LogFile) {
    try {
        Set-Content -Path $LogFile -Value $logHeader -Encoding UTF8 -ErrorAction Stop
    } catch {
        $script:LogEnabled = $false
    }
}

if ($script:LogEnabled -and $LogFile) {
    $journalDetail = $LogFile
    if ($script:LogFallback) {
        $journalDetail = "$LogFile (media script en lecture seule)"
    }
    Write-Status -Name "Journal" -Status INFO -Detail $journalDetail -NoCount
} else {
    Write-Status -Name "Journal" -Status INFO -Detail "desactive (aucun dossier accessible en ecriture)" -NoCount
}

Write-Section -Title "Prerequis" -Step "1/8"

$identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Status -Name "Droits admin" -Status FAIL -Detail "relancer via start.bat"
    Write-Summary
    Read-Host "  Entree pour quitter"
    exit 1
}
Write-Status -Name "Droits admin" -Status OK -NoCount
$script:Stats.OK = 0
Set-ProvConsoleMaximized
Write-Status -Name "Console" -Status INFO -Detail "fenetre maximisee" -NoCount

if (-not (Test-WingetAvailable)) {
    Write-Status -Name "winget" -Status FAIL -Detail "absent"
    Write-Host ""
    Write-Host ("  " + $UI.Arrow + " Installez 'Programme d'installation d'application' (App Installer)") -ForegroundColor Yellow
    Write-Host "    depuis le Microsoft Store, puis relancez." -ForegroundColor Yellow
    Write-Host "    ms-windows-store://pdp/?productid=9NBLGGH4NNS1" -ForegroundColor DarkGray
    Write-Summary
    Read-Host "  Entree pour quitter"
    exit 1
}
$ver = winget --version 2>$null
Write-Status -Name "winget" -Status OK -Detail "$ver" -NoCount
Initialize-ProvWingetNetwork

if (-not (Test-ProvOfflinePreflight)) {
    Write-Status -Name "Preflight" -Status FAIL -Detail "annule par l'operateur"
    Write-Summary
    Read-Host "  Entree pour quitter"
    exit 1
}

# Import fichiers ancien poste (avant installs)
Import-ProvMigrationUserFiles

Write-Section -Title "Identite du poste" -Step "2/8"
Set-ProvComputerIdentity
New-ProvLocalUser

Write-Section -Title "Applications (winget)" -Step "3/8"
Install-ChromeWithFallback | Out-Null
Import-ProvChromePasswordsFromMigration
foreach ($app in $script:AppsWinget) {
    Install-WingetApp -Name $app.Name -Id $app.Id | Out-Null
}

Write-Host ""
Install-AnyDeskDirect | Out-Null
Invoke-ProvNinite
Write-Host ""
Install-JavaRuntime

Install-ProvLobApps

Write-Section -Title "Microsoft 365 Apps (Business Premium)" -Step "4/8"
$officeOk = Install-Microsoft365

Write-Section -Title "Personnalisation du poste" -Step "5/8"
Set-ChromeAsDefault
Disable-ProvTaskbarWidgets
Set-ProvPerformanceTweaks
Set-ProvStartupApps
Set-ProvAnyDeskPassword
Set-ProvTaskbar
Remove-ProvBloatware
Set-ProvNetworkResources
Set-ProvDesktopOrganize

Write-Section -Title "Connexion Microsoft 365" -Step "6/8"
if ($officeOk) {
    Start-OfficeSignIn
} else {
    Write-Status -Name "Connexion Office" -Status FAIL -Detail ("suite non installee " + $UI.EmDash + " etape ignoree")
}

Write-Section -Title "Mises a jour & conformite" -Step "7/8"
Start-ProvWindowsUpdate
Test-ProvIntuneCompliance | Out-Null

Write-Section -Title "Inventaire & rapport" -Step "8/8"
$export = Export-ProvInventory
Sync-ProvReportsToUsb -Export $export | Out-Null
Send-ProvReport -Export $export

Write-Summary
Write-Log "Fin. OK=$($script:Stats.OK) SKIP=$($script:Stats.Skip) FAIL=$($script:Stats.Fail)" -Level INFO
if ($script:ReportPaths -and $script:ReportPaths.Count -gt 0) {
    Write-Host ("  " + $UI.Arrow + " Rapport : $($script:ReportPaths[0])") -ForegroundColor Gray
}
if ($script:ProvReportDir) {
    Write-Host ("  " + $UI.Arrow + " Dossier rapports (poste) : $($script:ProvReportDir)") -ForegroundColor Cyan
}
if ($script:ProvUsbReportDir) {
    Write-Host ("  " + $UI.Arrow + " Miroir cle USB : $($script:ProvUsbReportDir)") -ForegroundColor Cyan
}
Write-Host ("  " + $UI.Arrow + " Si le PC a ete renomme : redemarrer pour appliquer le nom.") -ForegroundColor Yellow
Write-Host ("  " + $UI.Arrow + " L'utilisateur finalise la connexion Office (mot de passe).") -ForegroundColor Gray
Write-Host ""

# Deploy termine : ne plus proposer la reprise au prochain run
Clear-ProvDeployState

# Nettoyage EN TOUT DERNIER (apres mail + resume) — confirmation OUI obligatoire
Invoke-ProvPostDeployCleanup
Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "  ERREUR FATALE :" -ForegroundColor Red
    Write-Host ("  " + $_.Exception.Message) -ForegroundColor Red
    try { Write-Log "FATAL: $($_.Exception.Message)" -Level ERR } catch {}
    Write-Host ""
}
finally {
    Read-Host "  Entree pour fermer"
}
