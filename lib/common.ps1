# ============================================================================
#  Provisioning — fonctions communes (dot-source depuis chaque script)
#  Fichier : lib\common.ps1
#
#  Import-ProvConfig : charge config.local.psd1 (sinon config.example.psd1 + WARN)
#  Get-ProvHostnameRegex / Get-ProvHostnameExamples : convention de nommage
#
#  UTF-8 BOM requis (Windows PowerShell 5.1).
# ============================================================================

function Import-ProvConfig {
    <#
    Cherche config.local.psd1 dans $BaseDir (racine du kit), sinon config.example.psd1.
    Avec l'exemple : avertissement, et IsExample = $true (renommage + envoi du rapport desactives).
    Retourne une hashtable, ou $null si aucune config lisible.
    #>
    param([Parameter(Mandatory)][string]$BaseDir)

    $local   = Join-Path $BaseDir "config.local.psd1"
    $example = Join-Path $BaseDir "config.example.psd1"
    $path = $null
    $isExample = $false
    if (Test-Path -LiteralPath $local) {
        $path = $local
    } elseif (Test-Path -LiteralPath $example) {
        $path = $example
        $isExample = $true
    } else {
        Write-Host "  [ERREUR] config.local.psd1 introuvable dans $BaseDir" -ForegroundColor Red
        return $null
    }

    try {
        $cfg = Import-PowerShellDataFile -LiteralPath $path -ErrorAction Stop
    } catch {
        Write-Host "  [ERREUR] config illisible ($path) : $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }

    # Valeurs par defaut pour les cles optionnelles
    $defaults = @{
        OrgName          = "Provisioning"
        FilePrefix       = "prov"
        WorkFolderName   = "Provisioning"
        ReportFolderName = "provisioning_reports"
        UsbReportsFolder = "reports"
        NoteFileName     = "README.txt"
        LocalAppRoot     = "C:\Provisioning"
        StartMenuFolder  = "Provisioning"
        UserAgent        = "windows-usb-provisioning"
        ReportTo         = @()
        SmtpServer       = "smtp.office365.com"
        SmtpPort         = 587
        SmtpUser         = ""
        SmtpEnabled      = $false
        NetworkDrives    = @()
        Printers         = @()
        Sites            = @()
        ItemTypes        = @()
        WingetApps       = @()
        NiniteFileNames  = @("Ninite.exe")
        OfflineMarkers   = @("README.txt", "AnyDesk.exe", "officedeploymenttool.exe", "odt.exe", "printers")
        PortableApp      = $null
        LobApp           = $null
        ClickOnceApp     = $null
        PrinterDriverNames   = @()
        PrinterDriverLabel   = ""
        PrinterDriverFolders = @("printers")
        PrinterInfExcludeDirPattern  = ""
        PrinterInfExcludeNamePattern = ""
        CameraClientEnabled           = $false
        CameraClientName              = ""
        CameraClientWingetName        = ""
        CameraClientWingetId          = ""
        CameraClientOfflineZip        = ""
        CameraClientOfflineSubdir     = ""
        CameraClientFilePattern       = ""
        CameraClientPreflightPattern  = ""
        CameraClientExePattern        = ""
        CameraClientInstalledPaths    = @()
        CameraClientProgramDirPattern = ""
        CameraClientExeNamePattern    = ""
        CameraClientArchiveUrl        = ""
        CameraClientArchiveVersion    = ""
        CameraClientCacheZipFormat    = ""
        CameraClientDownloadZipFormat = ""
    }
    foreach ($k in $defaults.Keys) {
        if (-not $cfg.ContainsKey($k)) { $cfg[$k] = $defaults[$k] }
    }
    $cfg["IsExample"]  = $isExample
    $cfg["ConfigPath"] = $path

    if ($isExample) {
        Write-Host ""
        Write-Host "  [WARN] config.local.psd1 absent — valeurs FICTIVES de config.example.psd1." -ForegroundColor Yellow
        Write-Host "         Copier config.example.psd1 en config.local.psd1 et l'adapter." -ForegroundColor Yellow
        Write-Host "         Renommage du PC et envoi du rapport desactives." -ForegroundColor Yellow
        Write-Host ""
    }

    foreach ($w in @(Test-ProvConfig -Config $cfg)) {
        Write-Host "  [WARN] config : $w" -ForegroundColor Yellow
    }
    return $cfg
}

function Test-ProvConfig {
    <# Retourne la liste des problemes detectes (vide = OK). #>
    param([Parameter(Mandatory)][hashtable]$Config)
    $issues = @()
    foreach ($k in @("OrgName", "OrgPrefix", "EmailDomain")) {
        if ([string]::IsNullOrWhiteSpace([string]$Config[$k])) { $issues += "$k vide" }
    }
    if ($Config.EmailDomain -and ([string]$Config.EmailDomain -notmatch '^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$')) {
        $issues += "EmailDomain invalide"
    }
    foreach ($r in @($Config.ReportTo)) {
        if ($r -and ([string]$r -notmatch '^[^@\s]+@[^@\s]+$')) { $issues += "ReportTo : adresse invalide" }
    }
    $sites = @(@($Config.Sites) | Where-Object { $_ -and $_.Code })
    $items = @(@($Config.ItemTypes) | Where-Object { $_ -and $_.Code })
    if ($sites.Count -eq 0) { $issues += "Sites vide" }
    if ($items.Count -eq 0) { $issues += "ItemTypes vide" }
    if ($sites.Count -gt 0 -and $items.Count -gt 0) {
        # NetBIOS : 15 caracteres max — PREFIX-SITE-ITEM-NNN
        $maxSite = ($sites | ForEach-Object { ([string]$_.Code).Length } | Measure-Object -Maximum).Maximum
        $maxItem = ($items | ForEach-Object { ([string]$_.Code).Length } | Measure-Object -Maximum).Maximum
        $len = ([string]$Config.OrgPrefix).Length + 1 + $maxSite + 1 + $maxItem + 1 + 3
        if ($len -gt 15) { $issues += "nom de poste le plus long = $len car. (> 15, limite NetBIOS)" }
    }
    return $issues
}

function Get-ProvHostnameRegex {
    <# ^PREFIX-(SITE1|SITE2...)-(ITEM1|ITEM2...)-\d{2,3}$ construit depuis la config. #>
    param([Parameter(Mandatory)][hashtable]$Config)
    $sites = @(@($Config.Sites) | Where-Object { $_ -and $_.Code } | ForEach-Object { [regex]::Escape([string]$_.Code) })
    $items = @(@($Config.ItemTypes) | Where-Object { $_ -and $_.Code } | ForEach-Object { [regex]::Escape([string]$_.Code) })
    return ('^{0}-({1})-({2})-\d{{2,3}}$' -f [regex]::Escape([string]$Config.OrgPrefix), ($sites -join '|'), ($items -join '|'))
}

function Get-ProvHostnameExamples {
    <# Deux exemples de noms valides (affichage). #>
    param([Parameter(Mandatory)][hashtable]$Config)
    $sites = @(@($Config.Sites) | Where-Object { $_ -and $_.Code })
    $items = @(@($Config.ItemTypes) | Where-Object { $_ -and $_.Code })
    if ($sites.Count -eq 0 -or $items.Count -eq 0) { return @("$($Config.OrgPrefix)-SITE-ITEM-001") }
    $s1 = if ($sites.Count -gt 1) { $sites[1].Code } else { $sites[0].Code }
    $i1 = if ($items.Count -gt 1) { $items[1].Code } else { $items[0].Code }
    return @(
        ("{0}-{1}-{2}-004" -f $Config.OrgPrefix, $s1, $items[0].Code),
        ("{0}-{1}-{2}-001" -f $Config.OrgPrefix, $sites[0].Code, $i1)
    )
}
