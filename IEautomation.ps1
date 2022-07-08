Import-Module Transferetto -Force
Import-Module ImportExcel -Force

$url = "website"
$username = 'username'
$password = 'password'
$hostname = 'ftpHost'
$sftp_username = 'sftpUsername'
$sftp_password = 'sftpPassword'
$file = 'C:\OutputPath\ExportFile.xlsm'
$outputFile = 'Extract.csv'
$logfilepath="C:\OutputPath\IEautomation.log"

Start-Sleep -Seconds 60

Start-Transcript -Path $logfilepath -Append

Write-Host -ForegroundColor Red "Removing existing files.";
Remove-Item $file -Force -ErrorAction SilentlyContinue
Remove-Item "C:\OutputPath\$outputFile" -Force -ErrorAction SilentlyContinue

# Create an ie com object
$ie = New-Object -com internetexplorer.application;
$ie.visible = $true;
$ie.navigate($url);
# Wait for the page to load
while ($ie.Busy -eq $true){ Start-Sleep -Milliseconds 1000; }

try
{
    # First login
    Write-Host -ForegroundColor Green "Attempting to login.";
    $ie.Document.getElementsByName("sample.login.user").item().value = $username
    $ie.Document.getElementsByName("sample.login.pass").item().value = $password
    $ie.Document.getElementsByName("sample.login.submit").item().click()   
}
catch
{
    Write-Host $_.Exception.Message
    Write-Host $_.InvocationInfo.ScriptLineNumber
    Stop-Transcript

    $client = Connect-SFTP -Server $hostname -Username $sftp_username -Password $sftp_password
    Add-SFTPFile -SftpClient $client -LocalPath $logfilepath -RemotePath "/IEautomation.log" -AllowOverride
    Disconnect-SFTP -SftpClient $client
}
try
{
    Do{Start-Sleep -Milliseconds 100}While($ie.Busy -eq $True)    
    Write-Host -ForegroundColor Green "Login successful.";

    #Select dropdown item
    Write-Host -ForegroundColor Green "Select dropdown item.";
    ($ie.Document.getElementById('placeHolder_export_DropDownList') | Where-Object { $_.innerHTML -eq 'WhatToSelect' }).selected = $true
    Do{Start-Sleep -Milliseconds 100}While($ie.Busy -eq $True)

    #Export til Excel
    Write-Host -ForegroundColor Green "Starting download export file.";
    $ie.Document.getElementById('placeHolder_export_Button').click();
    Do{Start-Sleep -Milliseconds 100}While($ie.Busy -eq $True)
   
    Write-Host -ForegroundColor Green "Saving file.";
    Start-Sleep -Seconds 1
    [void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")

    $ieProc = Get-Process | ? { $_.MainWindowHandle -eq $ie.HWND }
    [Microsoft.VisualBasic.Interaction]::AppActivate($ieProc.Id) #active IE window
    [System.Windows.Forms.SendKeys]::Sendwait("%{s}"); #press Ctrl+S to save file    
        
    Do
    {
        Start-Sleep -Seconds 1
        Write-Host -ForegroundColor Green "Downloading";
    }
    Until(Test-Path $file)
        
    Write-Host -ForegroundColor Green "Export file downloaded.";
    
    Start-Sleep -Seconds 10

    Write-Host -ForegroundColor Green "Quit Internet Explorer.";
    $ie.quit()

    Write-Host -ForegroundColor Green "Extract worksheet and export to CSV.";
    Import-Excel -Path $file -WorksheetName 'SheetToExtract' -NoHeader | Export-Csv -Path "C:\OutputPath\$outputFile" -Encoding UTF8 -NoTypeInformation -Force
    Write-Host -ForegroundColor Green "Worksheet exported to CSV.";
    
    Write-Host -ForegroundColor Green "Removing top row.";
    (Get-Content  "C:\OutputPath\$outputFile" | Select-Object -Skip 1) | Set-Content "C:\OutputPath\$outputFile" -Force


    Write-Host -ForegroundColor Green "Upload CSV to SFTP.";
    $secPassword = ConvertTo-SecureString $sftp_password -AsPlainText -Force
    $pscredential = New-Object -TypeName System.Management.Automation.PSCredential($sftp_username, $secPassword)

    $client = Connect-SFTP -Server $hostname -Username $sftp_username -Password $sftp_password
    Add-SFTPFile -SftpClient $client -LocalPath "C:\OutputPath\$outputFile" -RemotePath "/$outputFile" -AllowOverride -InformationAction SilentlyContinue | Out-Null
    Disconnect-SFTP -SftpClient $client
    Write-Host -ForegroundColor Green "Upload finished.";

    Stop-Transcript

    $client = Connect-SFTP -Server $hostname -Username $sftp_username -Password $sftp_password
    Add-SFTPFile -SftpClient $client -LocalPath $logfilepath -RemotePath "/IEautomation.log" -AllowOverride -InformationAction SilentlyContinue | Out-Null
    Disconnect-SFTP -SftpClient $client

    #Logoff
    shutdown /L /f
}
catch
{
    $_.Exception.Message
    $_.InvocationInfo.ScriptLineNumber
}
