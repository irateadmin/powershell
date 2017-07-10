#Function to get installed updates for the last 30 days, sort by InstalledOn and export the results to a CSV on the current users desktop

function Get-Updates { 

#You must have a "servers.txt" at "$env:USERPROFILE\Desktop\servers.txt" list of all the servers you want to scan for installed updates

$servers = Get-Content $env:USERPROFILE\Desktop\servers.txt    
$ErrorActionPreference = 'Stop' 
   
ForEach ($computer in $servers) {   
  
  try   
    {  
 
Get-HotFix -ComputerName $computer | Select-Object PSComputerName,HotFixID,InstalledOn,InstalledBy  | 
Where { $_.InstalledOn -gt (Get-Date).AddDays(-30) } |
sort InstalledOn |
Export-CSV $env:USERPROFILE\Desktop\Installed_Updates_Last_30_Days.csv
   
    }  
  
catch   
  
    {  
Add-content $computer -path "$env:USERPROFILE\Desktop\Unreachable_Servers.txt" 
    }   
}  
  
}

#Call Function
  
Get-Updates
