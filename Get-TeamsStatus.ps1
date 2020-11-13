﻿$headers = @{"Authorization"="Bearer <Insert token>";}
$Enable = 1
$CurrentStatus = "Offline"
DO {
# Get Teams Logfile and last icon overlay status
$TeamsStatus = Get-Content -Path "C:\Users\<UserName>\AppData\Roaming\Microsoft\Teams\logs.txt" -Tail 100 | Select-String -Pattern 'Setting the taskbar overlay icon - Available','Setting the taskbar overlay icon - Busy','Setting the taskbar overlay icon - Away','Setting the taskbar overlay icon - Do not disturb','Main window is closing','main window closed','Setting the taskbar overlay icon - On the phone','Setting the taskbar overlay icon - In a meeting','StatusIndicatorStateService: Added Busy','StatusIndicatorStateService: Added Available','StatusIndicatorStateService: Added InAMeeting','StatusIndicatorStateService: Added DoNotDisturb' | Select-Object -Last 1
# Get Teams Logfile and last app update deamon status
$TeamsActivity = Get-Content -Path "C:\Users\<UserName>\AppData\Roaming\Microsoft\Teams\logs.txt" -Tail 100 | Select-String -Pattern 'Resuming daemon App updates','Pausing daemon App updates' | Select-Object -Last 1

If ($TeamsStatus -like "*Setting the taskbar overlay icon - Available*" -or $TeamsStatus -like "*StatusIndicatorStateService: Added Available*") {
    $Status = "Available"
    Write-Host $Status
}
ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - Busy*" -or $TeamsStatus -like "*StatusIndicatorStateService: Added Busy*") {
    $Status = "Busy"
    Write-Host $Status
}
ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - Away*" -or $TeamsStatus -like "*StatusIndicatorStateService: Added Away*") {
    $Status = "Away"
    Write-Host $Status
}
ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - Do not disturb *" -or $TeamsStatus -like "*StatusIndicatorStateService: Added DoNotDisturb*") {
    $Status = "Do not disturb"
    Write-Host $Status
}
ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - In a meeting*" -or $TeamsStatus -like "*StatusIndicatorStateService: Added InAMeeting*") {
    $Status = "In a meeting"
    Write-Host $Status
}
ElseIf ($TeamsStatus -like "*ain window*") {
    $Status = "Offline"
    Write-Host $Status
}

If ($TeamsActivity -like "*Resuming daemon App updates*") {
    $Activity = "Not in a call"
    $ActivityIcon = "mdi:phone-off"
    Write-Host $Activity
}
ElseIf ($TeamsActivity -like "*Pausing daemon App updates*") {
    $Activity = "In a call"
    $ActivityIcon = "mdi:phone"
    Write-Host $Activity
}

If ($CurrentStatus -ne $Status) {
    $CurrentStatus = $Status

    $params = @{
     "state"="$CurrentStatus";
     "attributes"= @{
        "friendly_name"="Microsoft Teams status";
        "icon"="mdi:microsoft-teams";
        }
     }

    Invoke-RestMethod -Uri 'https://<HA URL>/api/states/sensor.teams_status' -Method POST -Headers $headers -Body ($params|ConvertTo-Json) -ContentType "application/json" 

}

If ($CurrentActivity -ne $Activity) {
    $CurrentActivity = $Activity

    $params = @{
     "state"="$Activity";
     "attributes"= @{
        "friendly_name"="Microsoft Teams activiteit";
        "icon"="$ActivityIcon";
        }
     }

    Invoke-RestMethod -Uri 'https://<HA URL>/api/states/sensor.teams_activity' -Method POST -Headers $headers -Body ($params|ConvertTo-Json) -ContentType "application/json" 

}
    Start-Sleep 1
} Until ($Enable -eq 0)