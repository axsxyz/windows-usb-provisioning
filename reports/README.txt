Reports mirror (USB stick)
==========================

After each full deploy, if the stick is writable (for example a BitLocker To Go
volume that has been unlocked), deploy-apps.ps1 copies the run output here:

  scripts\<UsbReportsFolder>\<PC-NAME>-<stamp>\
    rapport-<PC>-<stamp>.txt      deploy report
    inventaire-<PC>-<stamp>.csv   installed software inventory
    install-log-<stamp>.txt       full log
    <NoteFileName>                short note about the copy

The PC keeps its own copy in Documents\<ReportFolderName>\.
The mailto report to the IT recipients is still sent; this folder is a field archive.

These files contain machine and user data. Nothing in this folder is ever
committed: .gitignore excludes everything here except this file and .gitkeep.
The folder name on the stick is set by UsbReportsFolder in config.local.psd1
(default: reports).
