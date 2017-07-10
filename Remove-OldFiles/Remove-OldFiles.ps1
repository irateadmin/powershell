# fake reading in a list
#    in real life, use Get-Content
$SourceDirList = @"
c:\temp
d:\temp
"@.Split("`n").Trim()


$TimeStamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$ReportFile = "TooOldFiles_-_$TimeStamp.txt"
$ReportDir = $env:TEMP
$FullReportFile = Join-Path -Path $ReportDir -ChildPath $ReportFile

$MaxDaysOld = 30
$Today = Get-Date


$Results = foreach ($SourceDir in $SourceDirList)
    {
    # -Force gets hidden files and folders
    #    is it really neeeded here?
    $FileList = Get-ChildItem -Path $SourceDir -File -Recurse -Force
    foreach ($File in $FileList)
        {
        $DaysOld = ($Today - $File.LastAccessTime).Days
        if ($DaysOld -gt $MaxDaysOld)
            {
            # remove the -WhatIf to do it for real
            Remove-Item -LiteralPath $File.FullName -Force -WhatIf
            $Line = "{0,4}    {1}" -f $DaysOld, $File.FullName

            $Line
            }
        }
    }

$Results = (' Age    FileName', '----    --------') + $Results
$Results |
    Set-Content -Path $FullReportFile
