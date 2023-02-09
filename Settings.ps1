# Configure the variables below that will be used in the script

# Personal information to load correct languagepack
$locLang = "de" #en,nl

# Set MQTT Server setting
$mqttServerIP = "192.168.1.2"
$mqttUser = "mqttuser"
$mqttPass = "mqttpass"

# Set MQTT Topics
$toFirst = "MS-Teams"
$toSecond = "user"

$toLastConnect = $toFirst+"/"+$toSecond+"/lastconnect"
$toLastMessage = $toFirst+"/"+$toSecond+"/lastmassage"
$toStatus = $toFirst+"/"+$toSecond+"/status"
$toActivity = $toFirst+"/"+$toSecond+"/activity"
$toCam = $toFirst+"/"+$toSecond+"/camera"

# Set max line read from bottom of log file
# small: 250 , medium: 500, big: 1000
$ReadLines = 500 