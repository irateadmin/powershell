Import-module ActiveDirectory

if ($PSVersionTable.PSVersion -gt '6.0.0')
{

Function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton) {
    $browseForFolderOptions = 0
    if ($NoNewFolderButton) { $browseForFolderOptions += 512 }
    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
    return $selectedDirectory }
    
    # Show message box popup and return the button clicked by the user.
    Add-Type -AssemblyName System.Windows.Forms
    function Read-MessageBoxDialog([string]$Message, [string]$WindowTitle, [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::None) {  
    return [System.Windows.Forms.MessageBox]::Show($Message, $WindowTitle, $Buttons, $Icon) }
    
Read-MessageBoxDialog -Message "Select where you would like to save your report." -WindowTitle "Save Report Location" -Buttons OK
$OutFile = Read-FolderBrowserDialog -Message "Please select a directory" -InitialDirectory '%UserProfile%\Desktop\' -NoNewFolderButton
Read-MessageBoxDialog -Message "This may take several minutes to complete. Please wait for the 'Reports Complete' dialog at the end" -WindowTitle "Time Notification" -Buttons OK
$csvfile = $outfile + "\managerdirectreports.csv"
$managers = Get-ADUser -Filter * -Properties Name, DirectReports | Where-Object {($_.directreports -ne $Null) -and ($_.Enabled -eq 'True')}

foreach ($manager in $managers)
{
    $directreports = $manager.directreports 
    foreach($directreport in $directreports)
    {
        $user = Get-aduser $directreport -Properties name,samaccountname
        @{ Manager = $manager.Name
           DirectReport = $user.Name
        }| select Manager, DirectReport | Export-Csv $csvfile -NoTypeInformation -Append
    }
}
Read-MessageBoxDialog -Message "Reports Complete" -WindowTitle "Reports Complete" -Buttons OK
}

else {
    Write-Host "Please user Powershell v7 or greater to run this script"
}