<#
.NOTES
    Name: teamsstatus.ps1
    Author: heroprint
    Requires: PowerShell v2 or higher
    Fork: https://github.com/EBOOZ/TeamsStatus
    Github: https://github.com/heroprint/TeamsStatusMqtt
.SYNOPSIS
    Sets the status of the Microsoft Teams client via MQTT.
.DESCRIPTION
    This script is monitoring the Teams client logfile for certain changes. 
    The status status displays that availability  status of your Teams client 
    based on the icon overlay in the taskbar on Windows. 
    The activity status shows if you are in a call or not based on the App updates deamon, 
    which is paused as soon as you join a call.
    The cam status shows if you turn on your camera. 
.LINKS
    Teams Status Description: https://learn.microsoft.com/en-us/microsoftteams/presence-admins
    MDI Icons for Teams: https://icon-sets.iconify.design/mdi/microsoft-teams/
    Nice Fork for HA: https://github.com/AntoineGS/TeamsStatusV2/
    MQTT: https://github.com/eclipse/paho.mqtt.m2mqtt
    MQTT /w Powershell Example: https://jackgruber.github.io/2019-06-05-ps-mqtt/
#>

# Import settings, functions and language
.($PSScriptRoot + "\settings.ps1")
.($PSScriptRoot + "\functions.ps1")
.($PSScriptRoot + "\mqtt.ps1")
.($PSScriptRoot + "\$locLang.ps1")

# init 
$appDataFolder = GetAppDataFolder
$userName = GetUserName
$currentStatus = ""
$currentActivity = ""
$currentCam = ""

$teamsStatusHash = @{
    "Available" = $tsAvailable;
    "Busy" = $tsBusy;
    "Away" = $tsAway;
    "BeRightBack" = $tsBeRightBack;
    "DoNotDisturb" = $tsDoNotDisturb;
    "Offline" = $tsOffline;
    "Focusing" = $tsFocusing;
    "Presenting" = $tsPresenting;
    "InAMeeting" = $tsInAMeeting;
    "OnThePhone" = $tsOnThePhone;
}

MQTTConnect

# Start monitoring the Teams logfile when no parameter is used to run the script
Get-Content -Path "$appDataFolder\Microsoft\Teams\logs.txt" -Encoding Utf8 -Tail $ReadLines -ReadCount 0 -Wait | % {
    # Get Teams Logfile and last icon overlay status
    $TeamsStatus = $_ | Select-String -Pattern `
        'Setting the taskbar overlay icon -',`
        'StatusIndicatorStateService: Added' | Select-Object -Last 1

    # Get Teams Logfile and last app update deamon status
    $TeamsActivity = $_ | Select-String -Pattern `
        'Resuming daemon App updates',`
        'Pausing daemon App updates',`
        'SfB:TeamsNoCall',`
        'SfB:TeamsPendingCall',`
        'SfB:TeamsActiveCall',`
        'name: desktop_call_state_change_send, isOngoing',`
        'Attempting to play audio for notification type 1' | Select-Object -Last 1

    # Camstatus
    $CamStatus = $_ | Select-String -Pattern `
        'desktopClient createNativeRenderingResources',`
        'desktopClient destroyNativeRenderingResources' | Select-Object -Last 1

    # Get Teams application process
    $TeamsProcess = Get-Process -Name Teams -ErrorAction SilentlyContinue

    # Check if Teams is running and start monitoring the log if it is
    If ($TeamsProcess -ne $null) {

        #Check which is the last set stautus in log
        If($null -ne $TeamsStatus) {
            $teamsStatusHash.GetEnumerator() | ForEach-Object {
                If ($TeamsStatus -like "*Setting the taskbar overlay icon - $($_.value)*" -or `
                    $TeamsStatus -like "*StatusIndicatorStateService: Added $($_.key)*" -or `
                    $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: $($_.key) -> NewActivity*") {
                    $Status = $($_.value)
                }
            }
        }

        # Check the activity the is last set in log
        If($TeamsActivity -eq $null){ }
        ElseIf ($TeamsActivity -like "*Resuming daemon App updates*" -or `
            $TeamsActivity -like "*SfB:TeamsNoCall*" -or `
            $TeamsActivity -like "*name: desktop_call_state_change_send, isOngoing: false*") {
            $Activity = $taNotInACall
        }
        ElseIf ($TeamsActivity -like "*Pausing daemon App updates*" -or `
            $TeamsActivity -like "*SfB:TeamsActiveCall*" -or `
            $TeamsActivity -like "*name: desktop_call_state_change_send, isOngoing: true*") {
            $Activity = $taInACall
        }
        ElseIf ($TeamsActivity -like "*Attempting to play audio for notification type 1*") {
            $Activity = $taIncomingCall
        }
       
        # Check Camera
        If($CamStatus -eq $null){ }
        ElseIf ($CamStatus -like "*desktopClient createNativeRenderingResources*") {
            $Cam = $csCameraOn
        }
        ElseIf ($CamStatus -like "*desktopClient destroyNativeRenderingResources*") {
            $Cam = $csCameraOff
        }
    }
    # Set status to Offline when the Teams application is not running
    Else {
            $Status = $tsOffline
            $Activity = $taNotInACall
            $Cam = $csCameraOff
    }

    # Call MQTT API to set the status and activity 
    If ($CurrentStatus -ne $Status -and $Status -ne $null) {
        $CurrentStatus = $Status
        MQTTMsgSend $toStatus $Status
    }

    If ($CurrentActivity -ne $Activity -and $Activity -ne $null) {
        $CurrentActivity = $Activity
        MQTTMsgSend $toActivity $Activity
    }

    If ($CurrentCam -ne $Cam -and $Cam -ne $null) {
        $CurrentCam = $Cam
        MQTTMsgSend $toCam $Cam
    }
}
