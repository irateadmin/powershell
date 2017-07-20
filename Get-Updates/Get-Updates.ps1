#Get installed updates for the last 30 days, sort by InstalledOn and export the results to a CSV on the current users desktop
#You must have a "servers.txt" at "$env:USERPROFILE\Desktop\servers.txt" list of all the servers you want to scan for installed updates

$servers = Get-Content $env:USERPROFILE\Desktop\servers.txt    
$ErrorActionPreference = 'Stop' 

#Ask user for how many days of updates they want
$days = Read-Host "How many days back do you want to check for installed updates?"

#Store ForEach output in $Output   
$Output = ForEach ($server in $servers) {   
  
  try   
    {  
 
Get-HotFix -ComputerName $server | Select-Object PSComputerName,HotFixID,InstalledOn,InstalledBy  | 
Where { $_.InstalledOn -gt (Get-Date).AddDays(-$days) } | sort InstalledOn   
    }  
  
catch   
  
    {  
Add-content $server -path "$env:USERPROFILE\Desktop\Unreachable_Servers.txt" 
    }   
}  

#Write $Output to .csv 
$Output | Export-CSV $env:USERPROFILE\Desktop\Installed_Updates_Last_"$days"_Days.csv
