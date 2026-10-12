Offline folder (installers kept on the USB stick)
=================================================

Keep this folder ON THE USB STICK (for example X:\scripts\offline\).
Nothing here except this file and printers\SOURCE.txt is committed.

Search order used by deploy-apps.ps1:
  1) Removable drives : X:\scripts\offline , X:\offline , X:\usb_pro\downloads
  2) Other non-system drives (some sticks show up as fixed disks)
  3) The script folder (for example Desktop\scripts\offline)
  4) Desktop / Documents / Downloads

A folder only counts if it contains one of the OfflineMarkers from the config.

What can go here (the preflight treats the Office Deployment Tool and the
business-app folders declared in the config as critical; for the other items
the script falls back to winget or a direct download when they are missing):
  - AnyDesk.exe                       remote support client
  - officedeploymenttool.exe / odt.exe  Office Deployment Tool (Microsoft 365 Apps)
  - Ninite bundle                     file name from NiniteFileNames
  - Chrome offline installer          googlechromestandaloneenterprise64.msi
                                      or ChromeStandaloneSetup64.exe
  - Other offline installers          FirefoxSetup.exe, 7-Zip, Notepad++, ...
  - printers\                         printer driver package (see printers\SOURCE.txt)
  - <PortableApp.Folder>\             portable app, copied to <LocalAppRoot>
  - <LobApp.SourceFolders>\           business client folder, copied to <LocalAppRoot>
  - <ClickOnceApp.AppRefNames>        ClickOnce shortcut (network needed on first run)
  - camera client setup / zip         only if CameraClientEnabled = $true

Licences for anything placed here are the operator's responsibility.
