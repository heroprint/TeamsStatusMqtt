.($PSScriptRoot + "\settings.ps1")
.($PSScriptRoot + "\functions.ps1")

$appDataFolder = GetAppDataFolder

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
Unblock-File $PSScriptRoot\settings.ps1
Unblock-File $PSScriptRoot\teamstatus.ps1
Unblock-File $PSScriptRoot\functions.ps1
Unblock-File $PSScriptRoot\uninstall.ps1
Unblock-File $PSScriptRoot\de.ps1
Unblock-File $PSScriptRoot\en.ps1
Unblock-File $PSScriptRoot\nl.ps1

$TargetFile = $PSScriptRoot + "\TeamStatusMonitor.cmd"
$ShortcutFile = "$appDataFolder\Microsoft\Windows\Start Menu\Programs\Startup\TeamStatusMonitor.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $TargetFile
$Shortcut.WorkingDirectory = $PSScriptRoot
$Shortcut.Save()

Write-Output ""
Write-Output "Installation completed."
Write-Output "Please either reboot or launch the shortcut manually here:"
Write-Output "  $ShortcutFile"
Write-Output ""