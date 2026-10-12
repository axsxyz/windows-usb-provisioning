# config.example.psd1 — copy to config.local.psd1 (same folder) and edit.
# All values below are placeholders. config.local.psd1 is git-ignored.
# Save as UTF-8 with BOM (Windows PowerShell 5.1 misreads accents otherwise).
# Only literals are allowed here: strings, numbers, @(), @{}, $true, $false, $null.
@{
    # --- Organisation ---------------------------------------------------------
    OrgName          = 'Contoso'                    # banners, mail subject, report header, local-account description
    OrgPrefix        = 'CTS'                        # hostname prefix: CTS-<SITE>-<ITEM>-<NUM>
    FilePrefix       = 'prov'                       # generated files: prov-deploy-state.json, prov-migration.json, ...
    WorkFolderName   = 'Provisioning'               # %TEMP%, %LOCALAPPDATA%, %ProgramData% sub-folders
    ReportFolderName = 'provisioning_reports'       # under Documents (logs, inventory, report)
    UsbReportsFolder = 'reports'                    # mirror on the USB stick: scripts\reports\<PC>-<stamp>\
    NoteFileName     = 'README.txt'                 # name of the small notes the scripts drop next to their output
    LocalAppRoot     = 'C:\Contoso'                 # line-of-business apps are copied here
    StartMenuFolder  = 'Contoso'                    # Start Menu\Programs\<folder>
    UserAgent        = 'windows-usb-provisioning'   # sent to api.github.com when resolving latest releases

    # --- Identity / naming ----------------------------------------------------
    EmailDomain      = 'contoso.example'            # work email = first-name initial + surname @ domain (jdoe@contoso.example)
    Sites            = @(
        @{ Code = 'HQ';  Name = 'Head office' }
        @{ Code = 'NTH'; Name = 'North site' }
        @{ Code = 'STH'; Name = 'South site' }
    )
    ItemTypes        = @(
        @{ Code = 'PC';  Hint = '' }
        @{ Code = 'LAP'; Hint = 'laptop' }
        @{ Code = 'SRV'; Hint = '' }
        @{ Code = 'TAB'; Hint = 'tablette' }
    )

    # --- Report ---------------------------------------------------------------
    ReportTo         = @('it-reports@contoso.example')   # mailto recipients (one or more)
    SmtpServer       = 'smtp.office365.com'         # SMTP is not used (mailto only); kept for reference
    SmtpPort         = 587
    SmtpUser         = ''
    SmtpEnabled      = $false

    # --- Network resources ----------------------------------------------------
    NetworkDrives    = @(
        # @{ Letter = 'S'; Path = '\\fileserver.contoso.example\Shared' }
    )
    Printers         = @(
        # TCP/IP printer (driver staged from offline\printers\ via pnputil):
        @{ Name = 'Office MFP'; HostAddress = '192.0.2.10'; DriverName = 'Generic PCL6 Driver'; PortName = 'IP_192.0.2.10' }
        # Shared printer: @{ Name = 'Office MFP'; Path = '\\printserver.contoso.example\Office-MFP' }
    )
    # Driver package placed in offline\printers\ (see offline\printers\SOURCE.txt).
    PrinterDriverNames   = @('Generic PCL6 Driver')   # Windows driver names; the first one is the primary
    PrinterDriverLabel   = 'Generic PCL6'             # label shown in the console
    PrinterDriverFolders = @('printers\GenericPCL6', 'printers')   # searched in this order for *.inf
    PrinterInfExcludeDirPattern  = ''                 # regex on the INF full path to skip (e.g. a null/test driver folder)
    PrinterInfExcludeNamePattern = ''                 # regex on the INF file name to skip

    # --- Applications ---------------------------------------------------------
    WingetApps       = @(
        @{ Name = 'Mozilla Firefox'; Id = 'Mozilla.Firefox' }
        @{ Name = 'Discord';         Id = 'Discord.Discord' }
        @{ Name = '7-Zip';           Id = '7zip.7zip' }
        @{ Name = 'Bitwarden';       Id = 'Bitwarden.Bitwarden' }
        @{ Name = 'Notepad++';       Id = 'Notepad++.Notepad++' }
    )
    NiniteFileNames  = @('Ninite.exe')              # looked up in offline\ (any ninite*.exe also matches)

    # A folder is accepted as an "offline" asset folder if it contains one of these.
    OfflineMarkers   = @('README.txt', 'AnyDesk.exe', 'officedeploymenttool.exe', 'odt.exe', 'printers')

    # Portable app copied from offline\<Folder> to <LocalAppRoot>\<Folder> (+ Start Menu + Desktop shortcut).
    PortableApp      = $null
    # PortableApp    = @{ Name = 'Portable Tool'; Folder = 'PortableTool'; Exe = 'PortableTool.exe'
    #                     ShortcutAliases = @('Tool.lnk'); HeroPattern = '(?i)portable\s*tool' }

    # Java line-of-business client: folder copy, Start Menu shortcut, first run of its launcher JAR.
    LobApp           = $null
    # LobApp         = @{ Name = 'Business App'; Label = 'Business App'
    #                     SourceFolders = @('BusinessApp'); DestFolder = 'BusinessApp'
    #                     LauncherGlob = 'launcher*.jar'; KillAfterFirstRun = @('javaw', 'java', 'chrome')
    #                     Shortcut = 'Business App'; ShortcutAliases = @('BusinessApp.lnk')
    #                     HeroMatch = '(?i)business\s*app'; HeroExactPattern = '(?i)^business\s*app$'
    #                     HeroWebPattern = '(?i)business\s*app\s*web' }

    # ClickOnce application (needs network on first run).
    ClickOnceApp     = $null
    # ClickOnceApp   = @{ Name = 'Example Tool'; PreflightLabel = 'Example ClickOnce'
    #                     AppRefNames = @('Example Tool.appref-ms'); UrlFileName = 'Example-Tool.url.txt'
    #                     Url = 'https://vendor.example/tool/Example%20Tool.application'
    #                     InstalledFilter = '*Example*'; InstalledPattern = 'Example'
    #                     ShortcutNames = @('Example Tool.lnk') }

    # Optional camera / video-surveillance client chain:
    # offline setup or zip -> classic download URL -> winget fallback. Leave disabled if unused.
    CameraClientEnabled           = $false
    CameraClientName              = 'Camera Client'           # status label + Start Menu shortcut
    CameraClientWingetName        = 'Camera Client Lite'      # label of the winget fallback
    CameraClientWingetId          = ''                        # e.g. 'Vendor.CameraClientLite'
    CameraClientOfflineZip        = ''                        # glob, e.g. 'CameraClient*.zip' (also an offline marker)
    CameraClientOfflineSubdir     = ''                        # optional sub-folder of offline\ holding the setup
    CameraClientFilePattern       = ''                        # regex on setup file names in offline\
    CameraClientPreflightPattern  = ''                        # regex used by the preflight check
    CameraClientExePattern        = ''                        # regex on the setup exe inside a zip
    CameraClientInstalledPaths    = @()                       # paths relative to Program Files, e.g. 'Vendor\Client\Client.exe'
    CameraClientProgramDirPattern = ''                        # regex on Program Files folder names
    CameraClientExeNamePattern    = ''                        # regex on the installed exe name
    CameraClientClassicUrl        = ''                        # direct download of a zip (optional)
    CameraClientClassicVersion    = ''
    CameraClientCacheZipFormat    = ''                        # cached zip name in offline\, {0} = version
    CameraClientDownloadZipFormat = ''                        # downloaded zip name, {0} = version
}
