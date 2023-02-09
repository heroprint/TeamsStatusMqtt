# TeamsStatusMonitor with MQTT 
Works with:
* ioBroker
* openHab
* Home Assistant

## Introduction
This PowerShell script/service uses the local Teams' log file to track the status and activity of the logged in Teams user. Microsoft provides the status of your account via the Graph API, however to access the Graph API, your organization needs to grant consent for the organization so everybody can read their Teams status. This solution is great for anyone who's organization does not allow this.

This script makes use of three sensors that are published over mqtt to any home automation broker:
* teams_status
* teams_activity
* teams_camstatus

teams_status displays that availability status of your Teams client based on the icon overlay in the taskbar on Windows. 
teams_activity shows if you are in a call or not based on the App updates deamon, which is paused as soon as you join a call.
teams_camstatus shows if your camera is activ in a call


## Installation
* Download the files from this repository and save them to any folder (we will use C:\Scripts in this example)
* Configure the script in settings.ps1, open with Notepad++ or Win PowerShell ISE
* Start a elevated (Admin) PowerShell prompt, and execute the following scripts
  ```powershell
  Unblock-File C:\Scripts\install.ps1
  C:\Scripts\install.ps1
  ```
* Execute the file as requested in the Install.ps1 output
* After completing the steps above, start your Teams client and verify if the status and activity is updated as expected.
  
## Uninstallation
You can uninstall the service by executing the `uninstall.ps1` script.
Using the previous path as an example, in PowerShell you would run:
  ```powershell
  C:\Scripts\uninstall.ps1
  ```
Note: This will not stop the script if it is currently executing, if you would like to do so just kill it (powershell.exe).
If you get an error that the file "is not is not digitally signed", simply run the following before executing the uninstaller again:
  ```powershell
  Unblock-File C:\Scripts\uninstall.ps1
  ```

## Contribution
Pull Requests are welcomed!

## Credit
Original work by EBOOZ, which can be found here: https://github.com/EBOOZ/TeamsStatus.
Second inspiration work from AntoineGS, https://github.com/AntoineGS/TeamsStatusV2/.
But both works are only for Home Assistant, but i need for ioBroker a MQTT Interface. 

## Upcoming
Fixes and Impovments with error handling.
And Languages... please help me...


# Have fun with the script :)
Chris