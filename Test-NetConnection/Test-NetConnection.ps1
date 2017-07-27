#Check if OS version is Windows 7 (not supported) - Must use 8.1 or 2012 R2 or newer
$Version = (Get-WmiObject win32_operatingsystem).version

if ($version -lt "6.3.9600") {
    Write-Host "`nYou are using Windows $version which is not supported by this script. Please run this on Windows 8.1/Server 2012 R2 or newer"
    Read-Host "`nPress Enter to exit"
    Stop-Process -Id $PID
}

#Ask user what IP they want test
$IP = Read-Host "`nWhat IP address would you like to test?"

#Ask user if they want to test a specific port
$PortYesorNo = Read-Host "`nWould you like to test a specific port?"

#If Port is Yes ask for what Port number
$Port = Read-Host "`nWhat port would you like to test?"

#Run the test
Test-NetConnection $IP -port $Port

#Pause script so user can read output
Read-Host "`nPress Enter to exit"