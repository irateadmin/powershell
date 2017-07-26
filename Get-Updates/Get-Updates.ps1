#Get installed updates for the last userinput days, sort by InstalledOn and export the results to a CSV on the current users desktop
#You must have a "servers.txt" at "$env:USERPROFILE\Desktop\servers.txt" list of all the servers you want to scan for installed updates

#Ask user for how many days of updates they want
$days = Read-Host "How many days back do you want to check for installed updates?"

Function userinput {
$Output = ForEach ($server in $user_input_computer) {   
  
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
    } 

    function nouserinput {
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
    } 

$testpathservers = Test-Path $env:USERPROFILE\Desktop\servers.txt

If (-not $testpathservers) {
    $user_input_computer = Read-Host "What computer would you like to scan?"
    userinput    
    }
    else {
    $servers = Get-Content $env:USERPROFILE\Desktop\servers.txt
    nouserinput
    }