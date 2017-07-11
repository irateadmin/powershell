#Set the paths from the txt file
$SourceDirList = Get-Content C:\PowerShell\Remove-OldFiles\Paths_of_files_to_clean_up.txt


$TimeStamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$ReportFile = "TooOldFiles_-_$TimeStamp.txt"
$ReportDir = $env:TEMP
$FullReportFile = Join-Path -Path $ReportDir -ChildPath $ReportFile

$MaxDaysOld = 30
$Today = Get-Date


$Results = foreach ($SourceDir in $SourceDirList)
    {
    $FileList = Get-ChildItem -Path $SourceDir -File -Recurse
    foreach ($File in $FileList)
        {
        $DaysOld = ($Today - $File.LastAccessTime).Days
        if ($DaysOld -gt $MaxDaysOld)
            {
            #Remove the -WhatIf to do it for real
            Remove-Item -LiteralPath $File.FullName -Force -WhatIf
            $Line = "{0,4}    {1}" -f $DaysOld, $File.FullName

            $Line
            }
        }
    }

$Results = (' Age    FileName', '----    --------') + $Results
$Results |
    Set-Content -Path $FullReportFile