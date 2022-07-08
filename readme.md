# IE Automation custom tasks

Powershell script has the following functionality:

* Log script actions
* Login via IE
    * Select item from DropDownBox
    * Click download button
    * Save Excel(.xlms) export file in default IE download folder
* Quit IE
* Extract WorkSheet a work sheet and export to (.csv)
* Removes first row in the (.csv) file
* Upload (.csv) and (.log) file to SFTP site
* Logoff

## Prerequisites

Script has been tested successfully on Windoes 10 and Windows Server 2019. 

* Install Windows Feature '.NET Framework 3.5 (Includes .NET 2.0 and 3.0)'
* Set IE Internet Options Security to low (Windows 10) on Internet zone
* Uncheck 'Enable Protected Mode (requires restarting Internet Explorer)' in IE Internet Options Security (Windows Server 2019)
* Add website to Trusted sites in IE Internet Options Security
* Create folder 'C:\OutputPath' folder
* Download https://docs.microsoft.com/en-us/sysinternals/downloads/autologon and configure Autologon
* Run Windows Powershell as Administrator and install required modules
    * Install-Module Transferetto
    * Install-Module ImportExcel
    * Set-ExecutionPolicy -ExecutionPolicy Unrestricted
* Create Scheduled Task
    * Run only when user is logged on
    * Run with highest privileges
    * Configure for Windows 10 or Windows Server 2019
    * Triggers
        * At log on
        * Specific user: Insert current user here
    * Actions
        * Start a program
        * Program/script: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
        * Add arguments (optional): -file "C:\OutputPath\IEautomation.ps1"