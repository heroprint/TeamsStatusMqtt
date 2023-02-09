Add-Type -Path (Join-Path $PSScriptRoot "lib\M2Mqtt.Net.dll")

$MqttClient = [uPLibrary.Networking.M2Mqtt.MqttClient]($mqttServerIP)

function MQTTConnect(){
    $MqttClient.Connect([guid]::NewGuid(), $mqttUser,$mqttPass)
    $time = GetTimestamp
    MQTTMsgSend $toLastConnect $time
}

function MQTTDisconnect(){
    $MqttClient.Disconnect()
}

function MQTTMsgSend(){
    param (
        [Parameter(Mandatory = $true)] [string] $MqttTopic,
        [Parameter(Mandatory = $true)] [string] $MqttValue
    )
    $time = GetTimestamp
    $MqttClient.Publish( $MqttTopic, [System.Text.Encoding]::UTF8.GetBytes($MqttValue),0,0)
    $MqttClient.Publish( $toLastMessage , [System.Text.Encoding]::UTF8.GetBytes($time),0,0)
}
