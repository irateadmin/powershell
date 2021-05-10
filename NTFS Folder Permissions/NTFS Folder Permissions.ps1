#Hide Powershell Window
Add-Type -Name win -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);' -Namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle,0)

# Show an Open Folder Dialog and return the directory selected by the user.
Function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton) {
$browseForFolderOptions = 0
if ($NoNewFolderButton) { $browseForFolderOptions += 512 }
$app = New-Object -ComObject Shell.Application
$folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
return $selectedDirectory }

# Show input box popup and return the value entered by the user.
Add-Type -AssemblyName Microsoft.VisualBasic
function Read-InputBoxDialog([string]$Message, [string]$WindowTitle, [string]$DefaultText) {
return [Microsoft.VisualBasic.Interaction]::InputBox($Message, $WindowTitle, $DefaultText) }

# Show message box popup and return the button clicked by the user.
Add-Type -AssemblyName System.Windows.Forms
function Read-MessageBoxDialog([string]$Message, [string]$WindowTitle, [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::None) {  
return [System.Windows.Forms.MessageBox]::Show($Message, $WindowTitle, $Buttons, $Icon) }

Read-MessageBoxDialog -Message "Select where you would like to scan." -WindowTitle "Permissions Scan Location" -Buttons OK
$RootPath = Read-FolderBrowserDialog -Message "Please select a directory" -NoNewFolderButton

$maxdepth = Read-InputBoxDialog -Message "Please enter how many folders deep you want to scan" -WindowTitle "Folder Depth"

Read-MessageBoxDialog -Message "Select where you would like to save your report." -WindowTitle "Save Report Location" -Buttons OK
$OutFile = Read-FolderBrowserDialog -Message "Please select a directory" -InitialDirectory '%UserProfile%\Desktop\' -NoNewFolderButton

Read-MessageBoxDialog -Message "Your scan may take several minutes to complete. Please wait for the 'Scan Complete' dialog at the end" -WindowTitle "Time Notification" -Buttons OK

$Outfile = $OutFile + "\Permissions_Scan.csv"

$Header = "Folder Path,IdentityReference,AccessControlType,IsInherited,InheritanceFlags,PropagationFlags"

If (Test-Path $Outfile) {
Remove-Item $Outfile}  
$actual_depth_param = [int]([regex]::Matches($RootPath, "\\")).count + [int]$maxdepth + 1 
$Folders = dir $RootPath -recurse | where {$_.psiscontainer -eq $true}
Add-Content -Value $Header -Path $OutFile 
foreach ($Folder in $Folders){
 if (([regex]::Matches($Folder.fullname, "\\")).count -lt $actual_depth_param) { $ACLs = get-acl $Folder.fullname | ForEach-Object { $_.Access } 
 Foreach ($ACL in $ACLs){ 
 $OutInfo = $Folder.Fullname + "," + $ACL.IdentityReference + "," + $ACL.AccessControlType + "," + $ACL.IsInherited + "," + $ACL.InheritanceFlags + "," + $ACL.PropagationFlags  
 Add-Content -Value $OutInfo -Path $OutFile }}}
 
Read-MessageBoxDialog -Message "Scan Complete" -WindowTitle "Scan Complete" -Buttons OK