#####################################################################
#This Sample Code is provided for the purpose of illustration only
#and is not intended to be used in a production environment.  THIS
#SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT
#WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
#LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
#FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free
#right to use and modify the Sample Code and to reproduce and distribute
#the object code form of the Sample Code, provided that You agree:
#(i) to not use Our name, logo, or trademarks to market Your software
#product in which the Sample Code is embedded; (ii) to include a valid
#copyright notice on Your software product in which the Sample Code is
#embedded; and (iii) to indemnify, hold harmless, and defend Us and
#Our suppliers from and against any claims or lawsuits, including
#attorneys'' fees, that arise or result from the use or distribution
#of the Sample Code.
#####################################################################

# CUSTOMIZATIONS NEEDED:
# LINE 112 AND 113 ($From and $To)
# LINE 135 ($smtpServer)
# LINE 147 = Chose the appropriate drive letter for the backup location.

#Requires -Modules ADCSAdministration, Microsoft.PowerShell.Security,  PKI
#Requires -Version 3.0

#Region Help
<#

.SYNOPSIS Backs up Certificate Authority Issuing CA database and other
critical files

.DESCRIPTION This script utilizes native PowerShell cmdlets in Windows
Server 2012 R2 to perform daily backups for the CA database locally. After
those backups are completed, the static files should be backed up by the Enterprise backup solution or replicated off-box. Only 4 local backups are retained.

.OUTPUTS Daily e-mail report of backup status from previous day.

.EXAMPLE 
.\Backup-CertAuthority.ps1

#>
#####################################################################
#
# 
#####################################################################
#EndRegion

#Region ExecutionPolicy
#Set Execution Policy for Powershell
Set-ExecutionPolicy RemoteSigned
#EndRegion

#Region Modules
#Check IF required module is loaded, IF not load import it
IF (-not(Get-Module ADCSAdministration))
{
      Import-Module ADCSAdministration
}
If (-not(Get-Module PKI))
{
    Import-Module -Name PKI
}

IF (-not(Get-Module Microsoft.PowerShell.Security))
{
      Import-Module Microsoft.PowerShell.Security
}
#EndRegion

#Region Variables
#Dim variables
$Disk = Get-WmiObject -Class Win32_LogicalDisk -Namespace root\CIMv2 | Where-Object {$_.DeviceID -eq 'D:'} | Select-Object -Property DeviceID
$Limit = (Get-Date).AddDays(-1)
$myScriptName = $MyInvocation.MyCommand.Name
$evtProps = @("Index", "TimeWritten", "EntryType","Source", "InstanceID", "Message")
$LogonServer = $env:LOGONSERVER
$RetentionLimit = (Get-Date).AddDays(-3)
$ServerName = $env:COMPUTERNAME
#EndRegion

#Region Functions

Function fnGet-Date {#Begin function to get short date
      Get-Date -Format "MM-dd-yyyy"
}#End function fnGet-Date

Function fnGet-TodaysDate {#Begin function to get today's date
      Get-Date
}#End function fnGet-TodaysDate

Function fnGet-LongDate {#Begin function to get date and time in long format
      Get-Date -Format G
}#End function fnGet-LongDate

Function fnGet-ReportDate {#Begin function set report date format
      Get-Date -Format "yyyy-MM-dd"
}#End function fnGet-ReportDate

Function fnSend-AdminEmail {#Begin function to send summary e-mail to Administrators
      Param ($BodyMessage1, $BodyMessage2, $BodyMessage3, $BodyMessage4, $BodyMessage5, $BodyMessage6, $BodyMessage7, $BodyMessage8, $BodyMessage9)
      
      #Dim function specific variables
      $From = 'FromEmail@company.com'
      $To = "ToEmail@company.com"
      $Body = @"
            <p>$BodyMessage1</p>
            
            <p>$BodyMessage3</p>

            <p>$BodyMessage4</p>
            
            <p>$BodyMessage5</p>
            
            <p>$BodyMessage6</p>

        <p>$BodyMessage7</p>        

        <p>$BodyMessage8</p>
            
            <p>$BodyMessage9</p>
            
            <p>$BodyMessage2</p>
"@
      $ReportSubject="Active Directory Certificate Authority Backup status for $(fnGet-Date)"
      $smtpServer = 'SMTP Server'
      Send-MailMessage -From $From -To $To -Subject $ReportSubject -Body $Body -BodyAsHTML -SmtpServer $smtpServer
      }#End function fnSend-AdminEmail
#EndRegion

#Region Script
#Begin Script
$BodyMessage1 = "Script name: $myScriptName started at $(fnGet-LongDate)."

#Region Check folder structures
      $Disk = Get-WmiObject -Class Win32_LogicalDisk -Namespace root\CIMv2 | Where-Object {$_.DeviceID -eq 'C:'} | Select-Object -Property DeviceID
      $LogicalDisk = ($Disk).DeviceID
      #Subsitute folder name in $volBkpFldr with teh location where you will store the backups of your CA database
      $volBkpFldr = $LogicalDisk + "\CABackup\"
      $TodaysFldr = $volBkpFldr + $((Get-Date).ToString('yyyy-MM-dd'))
#EndRegion

#Region Backup-CA-DB
      #Backup Certificate Authority Database and Private Key
      If ((Test-Path -Path $TodaysFldr -PathType Container) -eq $true)
      {
            #Get password to be used during backup process
            #[String]$pw = fnGeneratePassword -length 25 -includeLowercaseLetters $true -includeUppercaseLetters $true -includeNumbers $true -includeSpecialChars $true
            #$BkpPassword = ConvertTo-SecureString -String $pw -AsPlainText -Force

            Backup-CARoleService -Path $TodaysFldr -DatabaseOnly
            $CAEvents = Get-EventLog -LogName Application -Newest 1 | Where-Object {$_.Source -eq "ESENT" -and $_.EventID -eq 213} | `
            Select-Object -Property TimeWritten, Source, InstanceID, Message | Sort-Object -Property TimeWritten
            If (($CAEvents).InstanceID -eq 213)
            {
                  $EventInfo = "Event Time: " + ($CAEvents).TimeWritten + " Event ID: " + ($CAEvents).InstanceID + " Event Details: " + ($CAEvents).Message
                  $BodyMessage3 = "Certificate Authority Backup Results:<br/>"
                  $BodyMessage3 += "Event Log Results:<br />"
                  $BodyMessage3 += $EventInfo  + "<br /><br />"
                  $BodyMessage2 = $pw
            }
            Else
            {
                  $BodyMessage3 = "Certificate Authority Backup Results:<br/>"
                  $BodyMessage3 += "Event Log Results:<br />"
                  $BodyMessage3 += "Certificate Authority Database and Private Key Backups failed $(fnGet-Date). Please investigate. <br />"
            }
      }
      Else
      {
            New-Item -ItemType Directory -Path $TodaysFldr -Force
            Backup-CARoleService -Path $TodaysFldr -DatabaseOnly
            $CAEvents = Get-EventLog -LogName Application -Newest 1 | Where-Object {$_.Source -eq "ESENT" -and $_.EventID -eq 213} | `
            Select-Object -Property TimeWritten, Source, InstanceID, Message | Sort-Object -Property TimeWritten
            If (($CAEvents).InstanceID -eq 213)
            {
                  $EventInfo = "Event Time: " + ($CAEvents).TimeWritten + " Event ID: " + ($CAEvents).InstanceID + " Event Details: " + ($CAEvents).Message
                  $BodyMessage3 = "Certificate Authority Backup Results:<br/>"
                  $BodyMessage3 += "Event Log Results:<br />"
                  $BodyMessage3 += $EventInfo  + "<br /><br />"
                  $BodyMessage2 = $pw
            }
            Else
            {
                  $BodyMessage3 = "Certificate Authority Backup Results:<br/>"
                  $BodyMessage3 += "Event Log Results:<br />"
                  $BodyMessage3 += "Certificate Authority Database and Private Key Backups failed $(fnGet-Date). Please investigate. <br />"
            }
      }
#EndRegion

#Region Backup-CA-Registry
#Backup Certificate Authority Registry Hive
      $RegFldr = $TodaysFldr + "\Registry"

      IF ((Test-Path -Path $RegFldr -PathType Container) -eq $true)
      {
            #Run reg.exe from command line to backup CA registry hive
            reg.exe export HKLM\System\CurrentControlSet\Services\CertSvc "$RegFldr\CARegistry_$(fnGet-ReportDate).reg"
            $RegFile = "$RegFldr\CARegistry_$(fnGet-ReportDate).reg"
            If ((Test-Path -Path $RegFile -PathType Leaf) -eq $true)
            {
                  $BodyMessage4 = "Registry Key Backup Results:<br />"
                  $BodyMessage4 += "Certificate Services registry key: $RegFile export was successful.<br />"
            }
            Else
            {
                  $BodyMessage4 = "Registry Key Backup Results:<br />"
                  $BodyMessage4 += "Certificate Services registry key export failed on $(fnGet-ReportDate).<br />"
            }
      }
      Else
      {
            New-Item -ItemType Directory -Path $RegFldr -Force
            reg.exe export HKLM\System\CurrentControlSet\Services\CertSvc  "$RegFldr\CARegistry_$(fnGet-ReportDate).reg"
            $RegFile = "$RegFldr\CARegistry_$(fnGet-ReportDate).reg"
            If ((Test-Path -Path $RegFile -PathType Leaf) -eq $true)
            {
                  $BodyMessage4 = "Registry Key Backup Results:<br />"
                  $BodyMessage4 += "Certificate Services registry key: $RegFile export was successful.<br />"
            }
            Else
            {
                  $BodyMessage4 = "Registry Key Backup Results:<br />"
                  $BodyMessage4 += "Certificate Services registry key export failed on $(fnGet-ReportDate).<br />"
            }     
      }
#EndRegion

#Region Backup-Policy-File
#If not using a Policy Certificate Authority server and policies are implemented using .INF file, backup configuration file.
#Backup Certificate Policy .Inf file
      $PolicyFldr = $TodaysFldr + "\PolicyFile\"
      $PolicyFile = $env:SystemRoot + "\CAPolicy.inf"
      IF ((Test-Path -Path $PolicyFldr -PathType Container) -eq $true)
      {
            Copy-Item -Path $PolicyFile -Destination $PolicyFldr
            If ((Test-Path -Path $PolicyFile -PathType Leaf) -eq $true)
            {
                  $BodyMessage5 = "Certificate Authority Policy File Backup Results:<br />"
                  $BodyMessage5 += "Backup copy of policy file: CAPolicy.inf was successful."
            }
            Else
            {
                  $BodyMessage5 = "Certificate Authority Policy File Backup Results:<br />"
                  $BodyMessage5 += "Backup copy of policy file: CAPolicy.inf failed. Please investigate."
            }
      }
      Else
      {
            New-Item -ItemType Directory -Path $PolicyFldr -Force
            Copy-Item -Path $PolicyFile -Destination $PolicyFldr
            If ((Test-Path -Path $PolicyFile -PathType Leaf) -eq $true)
            {
                  $BodyMessage5 = "Certificate Authority Policy File Backup Results:<br />"
                  $BodyMessage5 += "Backup copy of policy file: CAPolicy.inf was successful."
            }
            Else
            {
                  $BodyMessage5 = "Certificate Authority Policy File Backup Results:<br />"
                  $BodyMessage5 += "Backup copy of policy file: CAPolicy.inf failed. Please investigate."
            }
      }
#EndRegion

#Region Remove-Old-Backups
      #Cleanup old CA Backups
      $Results = Get-ChildItem -Path $volBkpFldr -Force | Where-Object {$_.LastWriteTime -le $RetentionLimit -and $_.PSisContainer} | Sort-Object -Property LastWriteTime -Descending
      $FolderName += ($Results).FullName
      $Results | Remove-Item -Force -Recurse
      If ($?) 
      {
            $BodyMessage8 = $FolderName + " was deleted as of $(fnGet-LongDate)"
      }
      Else 
      {
            $BodyMessage8 = "There are no backup folders ready for deletion as of $(fnGet-Date)."
      }     
#EndRegion

$BodyMessage9 = "Script name: $myScriptName completed at $(fnGet-LongDate)"

#Send E-mail Report
#fnSend-AdminEmail $BodyMessage1 $BodyMessage2 $BodyMessage3 $BodyMessage4 $BodyMessage5 $BodyMessage6 $BodyMessage7 $BodyMessage8 $BodyMessage9
#EndRegion
