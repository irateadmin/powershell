﻿#Function to get installed updates for the last 30 days, sort by InstalledOn and export the results to a CSV on the current users desktop

function Get-Updates {  
$servers = Get-Content $env:USERPROFILE\Desktop\servers.txt    
$ErrorActionPreference = 'Stop'    
ForEach ($computer in $servers) {   
  
  try   
    {  
 
Get-HotFix -cn $computer | Select-Object PSComputerName,HotFixID,InstalledOn,InstalledBy  | 
Where { $_.InstalledOn -gt (Get-Date).AddDays(-30) } |
sort InstalledOn | Export-CSV $env:USERPROFILE\Desktop\Installed_Updates_Last_30_Days.csv
   
    }  
  
catch   
  
    {  
Add-content $computer -path "$env:USERPROFILE\Desktop\Unreachable_Servers.txt" 
    }   
}  
  
}  
Get-Updates