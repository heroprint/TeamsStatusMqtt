<#
.NOTES
    Name: Get-TeamsStatus.ps1
    Author: Danny de Vries
    Requires: PowerShell v2 or higher
    Version History: https://github.com/EBOOZ/TeamsStatus/commits/main
.SYNOPSIS
    Sets the status of the Microsoft Teams client to Home Assistant.
.DESCRIPTION
    This script is monitoring the Teams client logfile for certain changes. It
    makes use of two sensors that are created in Home Assistant up front.
    The status entity (sensor.teams_status by default) displays that availability 
    status of your Teams client based on the icon overlay in the taskbar on Windows. 
    The activity entity (sensor.teams_activity by default) shows if you
    are in a call or not based on the App updates deamon, which is paused as soon as 
    you join a call.
.PARAMETER SetStatus
    Run the script with the SetStatus-parameter to set the status of Microsoft Teams
    directly from the commandline.
.EXAMPLE
    .\Get-TeamsStatus.ps1 -SetStatus "Offline"
#>
Add-Type -Path "C:\Scripts\lib\M2Mqtt.Net.dll"

# Import Settings
.($PSScriptRoot + "\Settings.ps1")

function MQTTMsgSend(){
    #SETUP MQTT CONNECTION
    $MqttClient = [uPLibrary.Networking.M2Mqtt.MqttClient]($mqttServerIP)
    $MqttClient.Connect([guid]::NewGuid(), $mqttUser,$mqttPass)
    if( $MqttClient.IsConnected ){
        $MqttClient.Publish("MS-Teams/"+$User+"/status", [System.Text.Encoding]::UTF8.GetBytes($Status), 0, 0)
        $MqttClient.Publish("MS-Teams/"+$User+"/activity", [System.Text.Encoding]::UTF8.GetBytes($Activity), 0, 0)
        $MqttClient.Disconnect()
    }
}

# Start monitoring the Teams logfile when no parameter is used to run the script
Get-Content -Path $env:APPDATA"\Microsoft\Teams\logs.txt" -Tail 1000 -ReadCount 0 -Wait | % {

    # Get Teams Logfile and last icon overlay status
    $TeamsStatus = $_ | Select-String -Pattern `
        'StatusIndicatorStateService: Added' | Select-Object -Last 1

    # Get Teams Logfile and last app update deamon status
    $TeamsActivity = $_ | Select-String -Pattern `
        'Resuming daemon App updates',`
        'Pausing daemon App updates',`
        'SfB:TeamsNoCall',`
        'SfB:TeamsPendingCall',`
        'SfB:TeamsActiveCall',`
        'name: desktop_call_state_change_send, isOngoing' | Select-Object -Last 1

    # Get Teams application process
    $TeamsProcess = Get-Process -Name Teams -ErrorAction SilentlyContinue

    # Check if Teams is running and start monitoring the log if it is
    If ($null -ne $TeamsProcess) {
        If($TeamsStatus -eq $null){ }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgAvailable*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Available*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Available -> NewActivity*") {
            $Status = $lgAvailable
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgBusy*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added Busy*" -or `
                $TeamsStatus -like "*Setting the taskbar overlay icon - $lgOnThePhone*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added OnThePhone*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Busy -> NewActivity*") {
            $Status = $lgBusy
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgAway*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added Away*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Away -> NewActivity*") {
            $Status = $lgAway
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgBeRightBack*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added BeRightBack*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: BeRightBack -> NewActivity*") {
            $Status = $lgBeRightBack
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgDoNotDisturb *" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added DoNotDisturb*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: DoNotDisturb -> NewActivity*") {
            $Status = $lgDoNotDisturb
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgFocusing*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added Focusing*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Focusing -> NewActivity*") {
            $Status = $lgFocusing
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgPresenting*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added Presenting*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Presenting -> NewActivity*") {
            $Status = $lgPresenting
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgInAMeeting*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added InAMeeting*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: InAMeeting -> NewActivity*") {
            $Status = $lgInAMeeting
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgOffline*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added Offline*") {
            $Status = $lgOffline
            Write-Host $Status
        }

        If($TeamsActivity -eq $null){ }
        ElseIf ($TeamsActivity -like "*Resuming daemon App updates*" -or `
            $TeamsActivity -like "*SfB:TeamsNoCall*" -or `
            $TeamsActivity -like "*name: desktop_call_state_change_send, isOngoing: false*") {
            $Activity = $lgNotInACall
            $ActivityIcon = $iconNotInACall
            Write-Host $Activity
        }
        ElseIf ($TeamsActivity -like "*Pausing daemon App updates*" -or `
            $TeamsActivity -like "*SfB:TeamsActiveCall*" -or `
            $TeamsActivity -like "*name: desktop_call_state_change_send, isOngoing: true*") {
            $Activity = $lgInACall
            $ActivityIcon = $iconInACall
            Write-Host $Activity
        }
    }
    # Set status to Offline when the Teams application is not running
    Else {
            $Status = $lgOffline
            $Activity = $lgNotInACall
            $ActivityIcon = $iconNotInACall
            Write-Host $Status
            Write-Host $Activity
    }

    # Call MQTT API to set the status and activity 
    If ($CurrentStatus -ne $Status) {
        $CurrentStatus = $Status
	 
        MQTTMsgSend
    }

    If ($CurrentActivity -ne $Activity) {
        $CurrentActivity = $Activity
 
       MQTTMsgSend
    }
}