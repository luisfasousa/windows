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

#Region Help
<#

.SYNOPSIS Backs up Certificate Authority Issuing CA database and other
critical files

.DESCRIPTION This script utilizes native PowerShell cmdlets in Windows
Server 2012 R2 to perform daily backups for the CA database locally. After
those backups are completed, the static files should be backed up by the Enterprise backup solution or replicated off-box. Only 1 local backup is retained.

EXAMPLE 
.\Backup-CertAuthority.ps1

#>
#EndRegion

$backupfolder = "D:\<CADataFolder>"

rd -Force -Recurse $backupfolder

sleep 5
certutil -backupdb $backupfolder

reg export HKLM\System\CurrentControlSet\Services\CertSvc\Configuration "$backupfolder\CAConfig.reg"

certutil -catemplates > "$backupfolder\Templates.txt"

net stop certsvc 
net start certsvc
sleep 5

