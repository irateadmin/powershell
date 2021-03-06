#Get installed updates for the last userinput days, sort by InstalledOn and export the results to a CSV on the current users desktop
#If you want to scan a list you should have "servers.txt" at "$env:USERPROFILE\Desktop\servers.txt" to scan for installed updates

#Ask user for how many days of updates they want
do {
    try {
        [Uint16]$days = Read-Host "`nHow many days back do you want to check for installed updates?"
        $inputOK = $true
    }
    catch {
       Write-Host "`nInvalid input! Please enter a numeric value." -ForegroundColor Red -BackgroundColor Black
    }
}
until ($inputOK)

#Function for if servers.txt does not exist
Function userinput {
    $Output = ForEach ($server in $user_input_computer) {   
        try {  
            Get-HotFix -ComputerName $server | Select-Object PSComputerName,HotFixID,InstalledOn,InstalledBy  | 
            Where { $_.InstalledOn -gt (Get-Date).AddDays(-$days) } | sort InstalledOn   
        }    
        catch {  
            Add-content $server -path "$env:USERPROFILE\Desktop\Unreachable_Machines.txt"
            Write-Host "Some machines were unreachable. The list is located here: '$env:USERPROFILE\Desktop\Unreachable_Machines.txt'`n" -ForegroundColor Red -BackgroundColor Black
        }   
    } 
    #Write $Output to .csv
    $Output | Export-CSV $env:USERPROFILE\Desktop\Installed_Updates_Last_"$days"_Days.csv
    Write-Host "Your scan is complete. The list is located here: '$env:USERPROFILE\Desktop\Installed_Updates_Last_$days`_Days.csv'" -ForegroundColor Green 
}

#Function for if servers.txt exists already 
function nouserinput {
    $Output = ForEach ($server in $servers) {  
        try {   
            Get-HotFix -ComputerName $server | Select-Object PSComputerName,HotFixID,InstalledOn,InstalledBy  | 
            Where { $_.InstalledOn -gt (Get-Date).AddDays(-$days) } | sort InstalledOn   
        }  
        catch {  
            Add-content $server -path "$env:USERPROFILE\Desktop\Unreachable_Machines.txt" 
            Write-Host "Some machines were unreachable. The list is located here: '$env:USERPROFILE\Desktop\Unreachable_Machines.txt'" -ForegroundColor Red -BackgroundColor Black
        }   
    }
    #Write $Output to .csv
    $Output | Export-CSV $env:USERPROFILE\Desktop\Installed_Updates_Last_"$days"_Days.csv
    Write-Host "Your scan is complete. The list is located here: '$env:USERPROFILE\Desktop\Installed_Updates_Last_$days`_Days.csv'" -ForegroundColor Green 
}

# Check if servers.txt exists and save result
$testpathservers = Test-Path $env:USERPROFILE\Desktop\servers.txt

# Check if Unreachable_Machines.txt already exists and remove it
$testpathunreachable_machines = Test-Path $env:USERPROFILE\Desktop\Unreachable_Machines.txt

#Remove Unreachable_Machines if it exists
If ($testpathunreachable_machines -eq $true) {
    Remove-Item $env:USERPROFILE\Desktop\Unreachable_Machines.txt
}

#Main Script - Calls Functions based on if servers.txt exists or not
If (-not $testpathservers) {
    Write-Host "`nA servers.txt file was not found here: '$env:USERPROFILE\Desktop\servers.txt'`n" -ForegroundColor Red -BackgroundColor Black
    Write-Host "You may enter a single computer name to scan.`n" -ForegroundColor Green
    $user_input_computer = Read-Host "What computer would you like to scan?"
    Write-Host "`nIt may take several minutes to complete your scan. Please be patient.`n" -ForegroundColor Green
    userinput  
}
else {
    $servers = Get-Content $env:USERPROFILE\Desktop\servers.txt
    Write-Host "A servers.txt file was found here: '$env:USERPROFILE\Desktop\servers.txt'`n" -ForegroundColor Green
    Write-Host "It may take several minutes to complete your scan. Please be patient.`n" -ForegroundColor Green
    nouserinput
}

#Pause script so user can read output from functions
Read-Host "`nPress Enter to exit"