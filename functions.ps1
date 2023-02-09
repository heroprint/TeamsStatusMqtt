function GetAppDataFolder {
    $userName = GetUserName
    return "C:\Users\$userName\AppData\Roaming"
}

function GetUserName {
	return $env:UserName
}

function GetTimestamp {
    return Get-Date -Format "dd.MM.yyyy HH:mm:ss"
}