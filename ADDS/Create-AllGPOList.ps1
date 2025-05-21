Import-Module activedirectory
Import-Module grouppolicy

# Create Temp folder if it doesn't exist
If (!(Test-Path -Path "c:\Temp" -ErrorAction SilentlyContinue )) {  New-Item "c:\Temp" -Type Directory -ErrorAction SilentlyContinue | Out-Null } 
 
(Get-ADForest).domains | 
  foreach { get-GPO -all -Domain $_ | 
    Select-Object @{n='Domain Name';e={$_.DomainName}}, @{n='GPO Name';e={$_.DisplayName}}, @{n='GPO Guid';e={$_.Id}} , @{n='Gpo Status';e={$_.GpoStatus}} , @{n='Creation Time';e={$_.CreationTime}} , @{n='Modification Time';e={$_.ModificationTime}} } | 
      Export-Csv c:\temp\AllGPOsList.csv
