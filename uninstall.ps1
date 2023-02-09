.($PSScriptRoot + "\functions.ps1")
$appDataFolder = GetAppDataFolder
$filename = "$appDataFolder\Microsoft\Windows\Start Menu\Programs\Startup\TeamStatusMonitor.lnk"

if (Test-Path $filename) {
  Remove-Item $filename
}