# ============================================================================
#  Provisioning — Rangement rapide du Bureau (standalone)
#  Fichier : desktop-organize.ps1
#  Usage   : menu start.bat  OU  powershell -File desktop-organize.ps1
#
#  Racine Bureau : Excel, Word, Chrome + apps metier de la config (LobApp, PortableApp)
#  _Inbox\Applications : autres apps
#  _Inbox\ : fichiers + dossiers restants (tri user)
#
#  UTF-8 BOM (Windows PowerShell 5.1)
# ============================================================================

param(
    [switch]$Yes,      # saute le prompt confirmation
    [switch]$NoPause   # pas de "Entree pour fermer" (appel depuis deploy)
)

$ErrorActionPreference = "Continue"
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    chcp 65001 | Out-Null
} catch {}

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
    Check = [string][char]0x2714
    Cross = [string][char]0x2716
    Info  = [string][char]0x2139
    Dot   = [string][char]0x00B7
    Play  = [string][char]0x25B8
    Arrow = [string][char]0x2192
    HLine = [string][char]0x2500
    TL    = [string][char]0x250C
    BL    = [string][char]0x2514
    VLine = [string][char]0x2502
}

function Write-DeskStatus {
    param([string]$Name, [ValidateSet('OK','FAIL','INFO','RUN','SKIP')][string]$Status, [string]$Detail = "")
    $icon = switch ($Status) {
        'OK'   { $UI.Check }
        'FAIL' { $UI.Cross }
        'INFO' { $UI.Info }
        'RUN'  { $UI.Play }
        'SKIP' { $UI.Dot }
    }
    $col = switch ($Status) {
        'OK'   { 'Green' }
        'FAIL' { 'Red' }
        'INFO' { 'DarkGray' }
        'RUN'  { 'Cyan' }
        'SKIP' { 'Yellow' }
    }
    $line = "  $icon  $Name"
    if ($Detail) { $line += "  —  $Detail" }
    Write-Host $line -ForegroundColor $col
}

function Read-ProvYesNo {
    param([Parameter(Mandatory)][string]$Prompt, [ValidateSet('Yes','No')][string]$Default = 'Yes')
    $hint = if ($Default -eq 'Yes') { 'O/n' } else { 'o/N' }
    while ($true) {
        $ans = (Read-Host "     $Prompt [$hint]").Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($ans)) { return ($Default -eq 'Yes') }
        if ($ans -in @('o','oui','y','yes','ok')) { return $true }
        if ($ans -in @('n','non','no')) { return $false }
        Write-Host "     Repondre O ou N." -ForegroundColor Yellow
    }
}

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

function Get-ProvDesktopHeroKind {
    <#
      Un seul heros par kind a la racine.
      Variante web de l'app metier / URL ≠ heros (→ Applications).
      Retourne : chrome | excel | word | portable | lob | $null
    #>
    param([string]$Name, [string]$Target = "")
    $n = "$Name"
    $t = "$Target"
    $blob = "$n $t"
    $lob = $ProvConfig.LobApp
    $pa = $ProvConfig.PortableApp

    # Variantes non-metier de l'app metier (web / URL) → pas heros
    if ($lob -and $lob.HeroMatch) {
        if (($lob.HeroWebPattern -and $n -match $lob.HeroWebPattern) -or $t -match '(?i)^https?://') {
            if ($n -match $lob.HeroMatch) { return $null }
        }
    }

    if ($blob -match '(?i)chrome\.exe' -or $n -match '(?i)^(google\s+)?chrome$') { return 'chrome' }
    if ($blob -match '(?i)excel\.exe' -or $n -match '(?i)^(microsoft\s+)?excel$') { return 'excel' }
    if ($blob -match '(?i)winword\.exe' -or $n -match '(?i)^(microsoft\s+)?word$') { return 'word' }
    if ($pa -and $pa.HeroPattern -and ($blob -match $pa.HeroPattern -or $n -match $pa.HeroPattern)) { return 'portable' }
    # App metier uniquement (pas Web)
    if ($lob -and $lob.HeroMatch) {
        if (($lob.HeroExactPattern -and $n -match $lob.HeroExactPattern) -or
            ($n -match $lob.HeroMatch -and $n -notmatch '(?i)web' -and $t -match '(?i)\.jar|javaw?\.exe|launcher')) {
            return 'lob'
        }
    }
    return $null
}

function Test-ProvDesktopHeroItem {
    param([string]$Name, [string]$Target = "")
    return [bool](Get-ProvDesktopHeroKind -Name $Name -Target $Target)
}

function Test-ProvDesktopAppShortcut {
    param([string]$Name, [string]$Target = "", [string]$Extension = "")
    if ($Extension -in @('.exe', '.msc')) { return $true }
    if ($Extension -notin @('.lnk', '.url', '.appref-ms')) { return $false }
    if (Test-ProvDesktopHeroItem -Name $Name -Target $Target) { return $false }
    return $true
}

function Get-ProvDesktopItemTarget {
    param($Item)
    if ($Item.PSIsContainer) { return "" }
    $ext = $Item.Extension.ToLowerInvariant()
    if ($ext -notin @('.lnk', '.url')) { return "" }
    try {
        $sh = New-Object -ComObject WScript.Shell
        return [string]$sh.CreateShortcut($Item.FullName).TargetPath
    } catch { return "" }
}

function Move-ProvDesktopItemUnique {
    param([string]$Source, [string]$DestDir)
    if (-not (Test-Path -LiteralPath $Source)) { return $false }
    if (-not (Test-Path -LiteralPath $DestDir)) {
        New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    }
    $name = Split-Path $Source -Leaf
    $dest = Join-Path $DestDir $name
    if (Test-Path -LiteralPath $dest) {
        $base = [IO.Path]::GetFileNameWithoutExtension($name)
        $ext = [IO.Path]::GetExtension($name)
        $i = 2
        do {
            $dest = Join-Path $DestDir ("{0}_{1}{2}" -f $base, $i, $ext)
            $i++
        } while (Test-Path -LiteralPath $dest)
    }
    try {
        Move-Item -LiteralPath $Source -Destination $dest -Force -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Test-ProvInboxFolderName {
    <# _Inbox, _Inbox (2), _Inbox_2, etc. #>
    param([string]$Name)
    return [bool]($Name -match '(?i)^_Inbox(\s*\(\d+\))?$' -or $Name -match '(?i)^_Inbox_\d+$')
}

function Get-ProvResolvedPath {
    param([string]$Path)
    if (-not $Path) { return "" }
    try { return [IO.Path]::GetFullPath($Path).TrimEnd('\') } catch { return $Path.TrimEnd('\') }
}

function Get-ProvDesktopRoots {
    <# Bureau user + Bureau public (Windows fusionne les deux a l'affichage). #>
    $list = New-Object System.Collections.Generic.List[string]
    $seen = @{}

    function Add-Root([string]$p) {
        if (-not $p -or -not (Test-Path -LiteralPath $p)) { return }
        $key = (Get-ProvResolvedPath $p).ToLowerInvariant()
        if (-not $key -or $seen.ContainsKey($key)) { return }
        $seen[$key] = $true
        [void]$list.Add((Get-ProvResolvedPath $p))
    }

    # Priorite : Bureau connu Windows (souvent deja OneDrive\Desktop)
    Add-Root ([Environment]::GetFolderPath("Desktop"))

    $public = Join-Path $env:PUBLIC "Desktop"
    Add-Root $public

    # Token eleve : profil interactif si different
    try {
        $ex = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" -EA SilentlyContinue).LastLoggedOnUser
        if ($ex -match '\\(.+)$') {
            $sam = $Matches[1]
            $prof = Join-Path "C:\Users" $sam
            Add-Root (Join-Path $prof "Desktop")
            Add-Root (Join-Path $prof "OneDrive\Desktop")
            # FR OneDrive
            Get-ChildItem -LiteralPath $prof -Directory -Filter "OneDrive*" -EA SilentlyContinue | ForEach-Object {
                Add-Root (Join-Path $_.FullName "Desktop")
                Add-Root (Join-Path $_.FullName "Bureau")
            }
            Add-Root (Join-Path $prof "Bureau")
        }
    } catch {}

    Write-Output -NoEnumerate ([string[]]$list.ToArray())
}

function Merge-ProvDesktopInboxes {
    <# Fusionne tous les _Inbox / _Inbox (2) vers un seul dossier principal. #>
    param(
        [Parameter(Mandatory)][string[]]$DesktopRoots,
        [Parameter(Mandatory)][string]$PrimaryInbox
    )
    $apps = Join-Path $PrimaryInbox "Applications"
    New-Item -ItemType Directory -Path $PrimaryInbox -Force | Out-Null
    New-Item -ItemType Directory -Path $apps -Force | Out-Null

    $inboxKey = (Get-ProvResolvedPath $PrimaryInbox).ToLowerInvariant()
    $merged = 0

    foreach ($desk in $DesktopRoots) {
        if (-not $desk) { continue }
        $dupes = @(Get-ChildItem -LiteralPath $desk -Force -Directory -EA SilentlyContinue |
            Where-Object { Test-ProvInboxFolderName -Name $_.Name })

        foreach ($dupe in $dupes) {
            $srcKey = (Get-ProvResolvedPath $dupe.FullName).ToLowerInvariant()
            if ($srcKey -eq $inboxKey) { continue }

            Write-DeskStatus -Name "_Inbox" -Status RUN -Detail ("fusion " + $dupe.FullName)

            $children = @(Get-ChildItem -LiteralPath $dupe.FullName -Force -EA SilentlyContinue)
            foreach ($child in $children) {
                if ($child.Name -eq $ProvConfig.NoteFileName) { continue }
                if ($child.PSIsContainer -and ($child.Name -eq 'Applications')) {
                    foreach ($appItem in @(Get-ChildItem -LiteralPath $child.FullName -Force -EA SilentlyContinue)) {
                        [void](Move-ProvDesktopItemUnique -Source $appItem.FullName -DestDir $apps)
                    }
                    try { Remove-Item -LiteralPath $child.FullName -Force -Recurse -EA SilentlyContinue } catch {}
                    continue
                }
                # Sous-inbox imbriquee → aplatir vers Applications ou racine
                if ($child.PSIsContainer -and (Test-ProvInboxFolderName -Name $child.Name)) {
                    foreach ($nested in @(Get-ChildItem -LiteralPath $child.FullName -Force -EA SilentlyContinue)) {
                        if ($nested.PSIsContainer -and $nested.Name -eq 'Applications') {
                            foreach ($appItem in @(Get-ChildItem -LiteralPath $nested.FullName -Force -EA SilentlyContinue)) {
                                [void](Move-ProvDesktopItemUnique -Source $appItem.FullName -DestDir $apps)
                            }
                        } else {
                            [void](Move-ProvDesktopItemUnique -Source $nested.FullName -DestDir $PrimaryInbox)
                        }
                    }
                    try { Remove-Item -LiteralPath $child.FullName -Force -Recurse -EA SilentlyContinue } catch {}
                    continue
                }
                [void](Move-ProvDesktopItemUnique -Source $child.FullName -DestDir $PrimaryInbox)
            }

            try {
                Remove-Item -LiteralPath $dupe.FullName -Force -Recurse -ErrorAction Stop
                $merged++
            } catch {
                # encore des fichiers verrouilles : laisser mais loguer
                Write-DeskStatus -Name "_Inbox" -Status SKIP -Detail ("reste " + $dupe.Name + " (verrouille ?)")
            }
        }
    }
    return $merged
}

function Ensure-ProvDesktopHeroShortcuts {
    param([string]$Desktop)
    $w = New-Object -ComObject WScript.Shell
    $heroes = @(
        @{
            Kind = 'chrome'; File = "Google Chrome.lnk"; Names = @("Google Chrome.lnk", "Chrome.lnk")
            Exes = @(
                "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
                "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
                "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
            )
        }
        @{
            Kind = 'excel'; File = "Excel.lnk"; Names = @("Excel.lnk", "Microsoft Excel.lnk")
            Exes = @(
                "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE"
            )
        }
        @{
            Kind = 'word'; File = "Word.lnk"; Names = @("Word.lnk", "Microsoft Word.lnk")
            Exes = @(
                "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE"
            )
        }
    )
    $pa = $ProvConfig.PortableApp
    if ($pa -and $pa.Folder -and $pa.Name) {
        $heroes += @{
            Kind = 'portable'; File = "$($pa.Name).lnk"; Names = @("$($pa.Name).lnk") + @($pa.ShortcutAliases | Where-Object { $_ })
            Exes = @(
                (Join-Path $ProvConfig.LocalAppRoot (Join-Path $pa.Folder (Join-Path $pa.Folder $pa.Exe))),
                (Join-Path $ProvConfig.LocalAppRoot (Join-Path $pa.Folder $pa.Exe))
            )
        }
    }
    $lob = $ProvConfig.LobApp
    if ($lob -and $lob.Shortcut) {
        $heroes += @{
            Kind = 'lob'; File = "$($lob.Shortcut).lnk"; Names = @("$($lob.Shortcut).lnk") + @($lob.ShortcutAliases | Where-Object { $_ })
            Exes = @(); StartMenu = @("$($lob.Shortcut).lnk")
        }
    }

    # Deja un heros de ce kind (n'importe quel nom) → ne pas recreer
    $existingKinds = @{}
    Get-ChildItem -LiteralPath $Desktop -Force -EA SilentlyContinue | Where-Object { -not $_.PSIsContainer } | ForEach-Object {
        $t = Get-ProvDesktopItemTarget -Item $_
        $k = Get-ProvDesktopHeroKind -Name $_.BaseName -Target $t
        if ($k) { $existingKinds[$k] = $true }
    }

    foreach ($h in $heroes) {
        if ($existingKinds.ContainsKey($h.Kind)) { continue }

        $srcLnk = Find-Shortcut -Names $h.Names
        if (-not $srcLnk -and $h.StartMenu) { $srcLnk = Find-Shortcut -Names $h.StartMenu }
        if ($srcLnk) {
            Copy-Item $srcLnk (Join-Path $Desktop $h.File) -Force -EA SilentlyContinue
            $existingKinds[$h.Kind] = $true
            continue
        }

        $exe = @($h.Exes) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
        if (-not $exe -and $h.Kind -eq 'portable') {
            $exe = Get-ChildItem (Join-Path $ProvConfig.LocalAppRoot $pa.Folder) -Filter $pa.Exe -Recurse -EA SilentlyContinue |
                Select-Object -First 1 -ExpandProperty FullName
        }
        if (-not $exe) {
            if ($h.Kind -eq 'lob') {
                $lobMenu = Join-Path $env:ProgramData (Join-Path (Join-Path "Microsoft\Windows\Start Menu\Programs" $ProvConfig.StartMenuFolder) "$($lob.Shortcut).lnk")
                if (Test-Path $lobMenu) {
                    Copy-Item $lobMenu (Join-Path $Desktop $h.File) -Force -EA SilentlyContinue
                    $existingKinds[$h.Kind] = $true
                }
            }
            continue
        }
        try {
            $s = $w.CreateShortcut((Join-Path $Desktop $h.File))
            $s.TargetPath = $exe
            $s.WorkingDirectory = Split-Path $exe -Parent
            $s.Save()
            $existingKinds[$h.Kind] = $true
        } catch {}
    }
}

function Invoke-ProvOrganizeOneDesktop {
    param(
        [Parameter(Mandatory)][string]$Desktop,
        [Parameter(Mandatory)][string]$InboxApps,
        [Parameter(Mandatory)][string]$InboxRoot,
        [Parameter(Mandatory)][hashtable]$HeroKindsSeen,
        [switch]$AllowHeroes,
        [switch]$EnsureHeroes
    )
    $movedApps = 0
    $movedInbox = 0
    $kept = 0
    $deduped = 0

    # Bureau public / secondaire : JAMAIS de heros (sinon 2e Chrome visible)
    if ($EnsureHeroes -and $AllowHeroes) {
        Ensure-ProvDesktopHeroShortcuts -Desktop $Desktop
    }

    $preferred = @{
        chrome   = 'Google Chrome.lnk'
        excel    = 'Excel.lnk'
        word     = 'Word.lnk'
        portable = $(if ($ProvConfig.PortableApp) { "$($ProvConfig.PortableApp.Name).lnk" } else { $null })
        lob      = $(if ($ProvConfig.LobApp) { "$($ProvConfig.LobApp.Shortcut).lnk" } else { $null })
    }

    $items = @(Get-ChildItem -LiteralPath $Desktop -Force -EA SilentlyContinue |
        Where-Object {
            $_.Name -notin @('desktop.ini', 'Thumbs.db', '.DS_Store') -and
            -not (Test-ProvInboxFolderName -Name $_.Name)
        })

    # 1) Heros : un seul au total (tous les Bureaux), uniquement sur le Bureau primaire
    $candidates = @($items | Where-Object { -not $_.PSIsContainer } | ForEach-Object {
        $t = Get-ProvDesktopItemTarget -Item $_
        $k = Get-ProvDesktopHeroKind -Name $_.BaseName -Target $t
        if (-not $k) { return }
        $prio = if ($_.Name -eq $preferred[$k]) { 0 } else { 1 }
        [pscustomobject]@{ Item = $_; Kind = $k; Prio = $prio }
    } | Sort-Object Prio, { $_.Item.Name })

    foreach ($c in $candidates) {
        $item = $c.Item
        if (-not (Test-Path -LiteralPath $item.FullName)) { continue }

        $already = $HeroKindsSeen.ContainsKey($c.Kind)
        if (-not $AllowHeroes -or $already) {
            if (Move-ProvDesktopItemUnique -Source $item.FullName -DestDir $InboxApps) { $deduped++ }
            continue
        }

        $HeroKindsSeen[$c.Kind] = $true
        $want = $preferred[$c.Kind]
        if ($want -and $item.Name -ne $want) {
            $dest = Join-Path $Desktop $want
            if (-not (Test-Path -LiteralPath $dest)) {
                try { Rename-Item -LiteralPath $item.FullName -NewName $want -EA Stop } catch {}
            }
        }
        $kept++
    }

    # 2) Relire puis ranger apps + dossiers/fichiers
    $items = @(Get-ChildItem -LiteralPath $Desktop -Force -EA SilentlyContinue |
        Where-Object {
            $_.Name -notin @('desktop.ini', 'Thumbs.db', '.DS_Store') -and
            -not (Test-ProvInboxFolderName -Name $_.Name)
        })

    foreach ($item in $items) {
        if (-not (Test-Path -LiteralPath $item.FullName)) { continue }
        $target = Get-ProvDesktopItemTarget -Item $item
        if (-not $item.PSIsContainer) {
            $kind = Get-ProvDesktopHeroKind -Name $item.BaseName -Target $target
            if ($kind) { continue }
        }

        $ext = if ($item.PSIsContainer) { "" } else { $item.Extension.ToLowerInvariant() }
        if (-not $item.PSIsContainer -and (Test-ProvDesktopAppShortcut -Name $item.BaseName -Target $target -Extension $ext)) {
            if (Move-ProvDesktopItemUnique -Source $item.FullName -DestDir $InboxApps) { $movedApps++ }
            continue
        }

        if ($Desktop -like "*\Public\Desktop" -and $item.PSIsContainer) { continue }

        if (Move-ProvDesktopItemUnique -Source $item.FullName -DestDir $InboxRoot) { $movedInbox++ }
    }

    return @{ Kept = $kept; Apps = $movedApps; Inbox = $movedInbox; Deduped = $deduped }
}

function Remove-ProvEmptyInboxDupes {
    <# Supprime les _Inbox vides / quasi-vides hors du principal. #>
    param(
        [Parameter(Mandatory)][string[]]$DesktopRoots,
        [Parameter(Mandatory)][string]$PrimaryInbox
    )
    $inboxKey = (Get-ProvResolvedPath $PrimaryInbox).ToLowerInvariant()
    $removed = 0
    $ignoreNames = @('desktop.ini', 'Thumbs.db', $ProvConfig.NoteFileName, '.DS_Store')

    foreach ($desk in $DesktopRoots) {
        if (-not $desk) { continue }
        foreach ($dupe in @(Get-ChildItem -LiteralPath $desk -Force -Directory -EA SilentlyContinue |
                Where-Object { Test-ProvInboxFolderName -Name $_.Name })) {
            $srcKey = (Get-ProvResolvedPath $dupe.FullName).ToLowerInvariant()
            if ($srcKey -eq $inboxKey) { continue }

            $kids = @(Get-ChildItem -LiteralPath $dupe.FullName -Force -EA SilentlyContinue |
                Where-Object { $_.Name -notin $ignoreNames })
            $isEmpty = ($kids.Count -eq 0)
            if (-not $isEmpty -and $kids.Count -eq 1 -and $kids[0].PSIsContainer -and $kids[0].Name -eq 'Applications') {
                $sub = @(Get-ChildItem -LiteralPath $kids[0].FullName -Force -EA SilentlyContinue |
                    Where-Object { $_.Name -notin $ignoreNames })
                $isEmpty = ($sub.Count -eq 0)
            }
            if (-not $isEmpty) { continue }

            try { cmd /c "attrib -R -S -H `"$($dupe.FullName)\*`" /S /D" | Out-Null } catch {}
            try {
                Remove-Item -LiteralPath $dupe.FullName -Force -Recurse -ErrorAction Stop
                $removed++
                Write-DeskStatus -Name "_Inbox" -Status OK -Detail ("supprime vide : " + $dupe.FullName)
            } catch {
                cmd /c "rd /s /q `"$($dupe.FullName)`"" | Out-Null
                if (-not (Test-Path -LiteralPath $dupe.FullName)) {
                    $removed++
                    Write-DeskStatus -Name "_Inbox" -Status OK -Detail ("supprime vide (rd) : " + $dupe.Name)
                } else {
                    Write-DeskStatus -Name "_Inbox" -Status SKIP -Detail ("vide verrouille : " + $dupe.Name)
                }
            }
        }
    }
    return $removed
}

function Set-ProvDesktopIconsCompact {
    <#
      One-shot : compacte les icones (plus de trous), puis desactive
      l'auto-arrange pour que l'utilisateur puisse reorganiser librement.
      Aligner sur la grille reste ON (pratique, sans figer l'ordre).
    #>
    $bag = "HKCU:\Software\Microsoft\Windows\Shell\Bags\1\Desktop"
    $ffOrig = [uint32]0
    try {
        if (-not (Test-Path -LiteralPath $bag)) {
            New-Item -Path $bag -Force | Out-Null
        }
        try {
            $cur = (Get-ItemProperty -LiteralPath $bag -Name FFlags -EA SilentlyContinue).FFlags
            if ($null -ne $cur) { $ffOrig = [uint32]$cur }
        } catch {}

        # 1) Activer auto-arrange + grille le temps du compactage
        $ffOn = $ffOrig -bor [uint32]0x1 -bor [uint32]0x8000
        Set-ItemProperty -LiteralPath $bag -Name FFlags -Type DWord -Value ([int64]$ffOn) -Force

        foreach ($n in @('IconLayouts', 'IconPos', 'IconStreams')) {
            Remove-ItemProperty -LiteralPath $bag -Name $n -EA SilentlyContinue
        }
    } catch {
        Write-DeskStatus -Name "Icones" -Status SKIP -Detail $_.Exception.Message
        return
    }

    function Invoke-ProvDesktopRefresh {
        try {
            $sig = @'
[DllImport("shell32.dll")] public static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
[DllImport("user32.dll")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
[DllImport("user32.dll")] public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
[DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
'@
            if (-not ("ProvDesk.Native" -as [type])) {
                Add-Type -MemberDefinition $sig -Name Native -Namespace ProvDesk -ErrorAction SilentlyContinue | Out-Null
            }
            [ProvDesk.Native]::SHChangeNotify(0x08000000, 0x1000, [IntPtr]::Zero, [IntPtr]::Zero)
            $progman = [ProvDesk.Native]::FindWindow("Progman", "Program Manager")
            if ($progman -ne [IntPtr]::Zero) {
                $shellDef = [ProvDesk.Native]::FindWindowEx($progman, [IntPtr]::Zero, "SHELLDLL_DefView", $null)
                if ($shellDef -eq [IntPtr]::Zero) {
                    $worker = [IntPtr]::Zero
                    do {
                        $worker = [ProvDesk.Native]::FindWindowEx([IntPtr]::Zero, $worker, "WorkerW", $null)
                        if ($worker -ne [IntPtr]::Zero) {
                            $shellDef = [ProvDesk.Native]::FindWindowEx($worker, [IntPtr]::Zero, "SHELLDLL_DefView", $null)
                        }
                    } while ($shellDef -eq [IntPtr]::Zero -and $worker -ne [IntPtr]::Zero)
                }
                if ($shellDef -ne [IntPtr]::Zero) {
                    [void][ProvDesk.Native]::PostMessage($shellDef, 0x100, [IntPtr]0x74, [IntPtr]::Zero)
                }
            }
        } catch {}
        try { & "$env:SystemRoot\System32\ie4uinit.exe" -show 2>$null | Out-Null } catch {}
    }

    Invoke-ProvDesktopRefresh
    Start-Sleep -Milliseconds 800

    # 2) Couper auto-arrange (bit 0x1) — garder align grille ; positions compactees restent
    try {
        $ffOff = ($ffOrig -bor [uint32]0x8000) -band (-bnot [uint32]0x1)
        # Si l'origine avait deja auto-arrange, on le force OFF quand meme (one-shot)
        $ffOff = $ffOff -band (-bnot [uint32]0x1)
        Set-ItemProperty -LiteralPath $bag -Name FFlags -Type DWord -Value ([int64]$ffOff) -Force
    } catch {}

    Invoke-ProvDesktopRefresh
    Write-DeskStatus -Name "Icones" -Status OK -Detail "compact one-shot (auto-arrange off — user libre)"
}

function Get-ProvHeroLabelList {
    $labels = @('Excel', 'Word', 'Chrome')
    if ($ProvConfig.LobApp -and $ProvConfig.LobApp.Shortcut) { $labels += [string]$ProvConfig.LobApp.Shortcut }
    if ($ProvConfig.PortableApp -and $ProvConfig.PortableApp.Name) { $labels += [string]$ProvConfig.PortableApp.Name }
    return ($labels -join ', ')
}

function Invoke-ProvDesktopOrganize {
    param([switch]$SkipPrompt)

    Write-Host ""
    Write-Host ("  " + $UI.TL + $UI.HLine + " Rangement Bureau " + ($UI.HLine * 34)) -ForegroundColor Magenta
    Write-Host ("  " + $UI.VLine + "  Racine  : " + (Get-ProvHeroLabelList)) -ForegroundColor Gray
    Write-Host ("  " + $UI.VLine + "  _Inbox\Applications : autres apps (+ doublons Public)") -ForegroundColor Gray
    Write-Host ("  " + $UI.VLine + "  _Inbox\ : fichiers + dossiers restants (tri user)") -ForegroundColor Gray
    Write-Host ("  " + $UI.VLine + "  Heros uniquement sur Bureau user — Public strippe") -ForegroundColor DarkGray
    Write-Host ("  " + $UI.BL + ($UI.HLine * 53)) -ForegroundColor Magenta
    Write-Host ""

    if (-not $SkipPrompt -and -not $Yes) {
        if (-not (Read-ProvYesNo -Prompt "Ranger le Bureau maintenant ?" -Default Yes)) {
            Write-DeskStatus -Name "Bureau" -Status SKIP -Detail "rangement ignore"
            return
        }
    }

    $roots = New-Object System.Collections.Generic.List[string]
    foreach ($r in @(Get-ProvDesktopRoots)) {
        if ($null -eq $r) { continue }
        if ($r -is [System.Array]) {
            foreach ($x in $r) {
                if ($x -and -not $roots.Contains([string]$x)) { [void]$roots.Add([string]$x) }
            }
        } else {
            $s = [string]$r
            if ($s -and -not $roots.Contains($s)) { [void]$roots.Add($s) }
        }
    }
    if ($roots.Count -eq 0) {
        Write-DeskStatus -Name "Bureau" -Status FAIL -Detail "chemin Bureau introuvable"
        return
    }

    $primary = [Environment]::GetFolderPath("Desktop")
    if (-not $primary -or -not (Test-Path -LiteralPath $primary)) {
        $primary = ($roots | Where-Object { $_ -notlike "*\Public\Desktop" } | Select-Object -First 1)
    }
    if (-not $primary) { $primary = $roots[0] }
    $primary = Get-ProvResolvedPath $primary
    $primaryKey = $primary.ToLowerInvariant()

    $inbox = Join-Path $primary "_Inbox"
    $apps  = Join-Path $inbox "Applications"
    New-Item -ItemType Directory -Path $apps -Force | Out-Null

    $merged = Merge-ProvDesktopInboxes -DesktopRoots @($roots.ToArray()) -PrimaryInbox $inbox
    if ($merged -gt 0) {
        Write-DeskStatus -Name "_Inbox" -Status OK -Detail ("$merged doublon(s) fusionne(s)")
    }
    [void](Remove-ProvEmptyInboxDupes -DesktopRoots @($roots.ToArray()) -PrimaryInbox $inbox)

    $heroSeen = @{}
    $totKept = 0; $totApps = 0; $totInbox = 0; $totDedup = 0

    # Primaire d'abord (heros), puis Public / secondaires (strip heros)
    $ordered = @($roots | Sort-Object {
        $k = (Get-ProvResolvedPath $_).ToLowerInvariant()
        if ($k -eq $primaryKey) { 0 }
        elseif ($_ -like "*\Public\Desktop") { 2 }
        else { 1 }
    })

    foreach ($desk in $ordered) {
        $deskRes = Get-ProvResolvedPath $desk
        $isPrimary = ($deskRes.ToLowerInvariant() -eq $primaryKey)
        Write-DeskStatus -Name "Scan" -Status RUN -Detail $deskRes
        $r = Invoke-ProvOrganizeOneDesktop `
            -Desktop $deskRes `
            -InboxApps $apps `
            -InboxRoot $inbox `
            -HeroKindsSeen $heroSeen `
            -AllowHeroes:$isPrimary `
            -EnsureHeroes:$isPrimary
        $totKept  += $r.Kept
        $totApps  += $r.Apps
        $totInbox += $r.Inbox
        $totDedup += $r.Deduped
    }

    $merged2 = Merge-ProvDesktopInboxes -DesktopRoots @($roots.ToArray()) -PrimaryInbox $inbox
    if ($merged2 -gt 0) {
        Write-DeskStatus -Name "_Inbox" -Status OK -Detail ("$merged2 doublon(s) fusionne(s) apres scan")
    }
    $emptied = Remove-ProvEmptyInboxDupes -DesktopRoots @($roots.ToArray()) -PrimaryInbox $inbox
    if ($emptied -gt 0) {
        Write-DeskStatus -Name "_Inbox" -Status OK -Detail ("$emptied vide(s) supprime(s)")
    }

    $note = @"
_Inbox — a trier
================
Applications\  = raccourcis apps + doublons (2e Chrome Public, Edge…)
Racine Inbox   = fichiers et dossiers recuperes du Bureau

Heros Bureau (1 seul chacun, Bureau user uniquement) :
$(Get-ProvHeroLabelList)
"@
    Set-Content -Path (Join-Path $inbox $ProvConfig.NoteFileName) -Value $note -Encoding UTF8 -Force

    Set-ProvDesktopIconsCompact

    Write-DeskStatus -Name "Bureau" -Status OK -Detail ("heros=$totKept  dedup=$totDedup  apps→Applications=$totApps  reste→Inbox=$totInbox")
    Write-Host ("  " + $UI.Arrow + " Bureau range + icones compactees.") -ForegroundColor Cyan
}

# --- MAIN (standalone) ---
try {
    if (-not $NoPause) {
        Clear-Host
        Write-Host ""
        Write-Host ("  " + ("=" * 58)) -ForegroundColor DarkCyan
        Write-Host ("   " + $ProvConfig.OrgName + "  ·  Rangement rapide du Bureau") -ForegroundColor Cyan
        Write-Host ("  " + ("=" * 58)) -ForegroundColor DarkCyan
    }
    Invoke-ProvDesktopOrganize
    Write-Host ""
} catch {
    Write-Host ("  ERREUR : " + $_.Exception.Message) -ForegroundColor Red
}
if (-not $NoPause) {
    Read-Host "  Entree pour fermer"
}
