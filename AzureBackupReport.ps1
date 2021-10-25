<#
.Synopsis
   creates a report about SCDPM Backup
.DESCRIPTION
   this script creates and send a report about the state from System Center DataProtection Manager 2012 R2
   - recovery points
   - jobs in progress
   - failed jobs for yesterday and today
   - agent status
   - disk status
.EXAMPLE
   DPM-Report.ps1
.NOTES
   Author:    Josh Burkard - mRLedV9cTF2k0GlpfxQa@burkard.it
   Date:      22.10.2015
   Version:   1.0.0
   Requires:  PowerShell V2
              Module yxz in directory C:\Windows\System32\WindowsPowerShell\v1.0\Modules (for x64)
              Module xyz in directory C:\Windows\SysWOW64\WindowsPowerShell\v1.0\Modules (for x86)
#>

#region ScriptStart
    $elapsed = [System.Diagnostics.Stopwatch]::StartNew()
    if ($Host.Version.Major -ne 2)
    {
        clear
        write-host "Started at $(get-date)"
    }

    # Set Error Action to Silently Continue
    # $ErrorActionPreference = "SilentlyContinue"
    # $WarningPreference = "SilentlyContinue"
#endregion ScriptStart

#region Declarations
    $sendEmail                = $true
    $saveHTML                 = $false
    $emailTo                  = '<techopscalendar@deltadentalia.com>'
    $emailFrom                = '<azurebackup@deltadentalia.com>'
    #$emailSubjectPrefix       = '[Information]:'
    $emailSubject             = 'Azure Backup Report'
    $emailHost                = 'internalrelay.deltadentalia.com'
    $file                     = "$($env:TEMP)\dpm-backup-report.html"

    $DPMServerName            = 'ddiaazbk01.deltadentalia.com'
    $tableWidth               = '1000'

    $colorGreen               = 'rgb(0,   176,  80)'
    $colorYellow              = 'rgb(255,  45, 198)'
    $colorOrange              = 'rgb(255, 128,   0)'
    $colorRed                 = 'rgb(251, 152, 149)'
    $colorGray                = 'rgb(228, 228, 228)'

    $showJobsInProgress       = $true
    $showJobsFailed           = $false
    $showAgentStatus          = $true
    $showAgentsDeleted        = $false
    $showDisks                = $true

    $RecoveryPointAgeOK       = 1 # days
    $RecoveryPointAgeWarning  = 7 # days
#endregion Declarations

#region Functions
Function ConvertTo-AdvHTML
    {
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [Object[]]$InputObject,
            [string[]]$HeadWidth,
            [string[]]$HeadwidthPercent,
            [string]$CSS = "", 
            [string]$CSS_old = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse; width: 1000px;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;font-size:120%;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
</style>
"@,
            [string]$Title,
            [string]$PreContent,
            [string]$PostContent,
            [string]$Body,
            [switch]$Fragment
        )
    
        Begin {
            If ($Title)
            {   $CSS += "`n<title>$Title</title>`n"
            }
            $Params = @{
                Head = $CSS
            }
            If ($PreContent)
            {   $Params.Add("PreContent",$PreContent)
            }
            If ($PostContent)
            {   $Params.Add("PostContent",$PostContent)
            }
            If ($Body)
            {   $Params.Add("Body",$Body)
            }
            If ($Fragment)
            {   $Params.Add("Fragment",$true)
            }
            $Data = @()
        }
    
        Process {
            ForEach ($Line in $InputObject)
            {   $Data += $Line
            }
        }
    
        End {
            $Html = $Data | ConvertTo-Html -Fragment

            $NewHTML = @()
            ForEach ($Line in $Html)
            {   
                If ($Line -like "*<table>*")
                {
                    $Line = $Line.Replace('<table>', '<table width="' + $tableWidth + '">')
                }
                If ($Line -like "*<th>*")
                {   
                    If ($Headwidth)
                    {
                        $Index = 0
                        $Reg = $Line | Select-String -AllMatches -Pattern "<th>(.*?)<\/th>"
                        ForEach ($th in $Reg.Matches)
                        {   
                            If ($Index -le ($HeadWidth.Count - 1))
                            {
                                If ($HeadWidth[$Index] -and $HeadWidth[$Index] -gt 0)
                                {
                                    $Line = $Line.Replace($th.Value,"<th width=""$($HeadWidth[$Index])"" style=""width:$($HeadWidth[$Index])px"">$($th.Groups[1])</th>")
                                }
                            }
                            $Index ++
                        }
                    }
                    If ($HeadwidthPercent)
                    {   $Index = 0
                        $Reg = $Line | Select-String -AllMatches -Pattern "<th>(.*?)<\/th>"
                        ForEach ($th in $Reg.Matches)
                        {   
                            If ($Index -le ($HeadwidthPercent.Count - 1))
                            {
                                If ($HeadwidthPercent[$Index] -and $HeadwidthPercent[$Index] -gt 0)
                                {
                                    $Line = $Line.Replace($th.Value,"<th style=""width:$($HeadwidthPercent[$Index])%"">$($th.Groups[1])</th>")
                                }
                            }
                            $Index ++
                        }
                    }
                }
        
                Do {
                    Switch -regex ($Line)
                    {   "<td>\[cell:(.*?)\].*?<\/td>"
                        {
                            $Line = $Line.Replace("<td>[cell:$($Matches[1])]","<td style=""background-color:$($Matches[1])"">")
                            Break
                        }
                        "\[cellclass:(.*?)\]"
                        {   $Line = $Line.Replace("<td>[cellclass:$($Matches[1])]","<td class=""$($Matches[1])"">")
                            Break
                        }
                        "\[row:(.*?)\]"
                        {   $Line = $Line.Replace("<tr>","<tr style=""background-color:$($Matches[1])"">")
                            $Line = $Line.Replace("[row:$($Matches[1])]","")
                            Break
                        }
                        "\[rowclass:(.*?)\]"
                        {   $Line = $Line.Replace("<tr>","<tr class=""$($Matches[1])"">")
                            $Line = $Line.Replace("[rowclass:$($Matches[1])]","")
                            Break
                        }
                        "<td>\[bar:(.*?)\](.*?)<\/td>"
                        {   $Bar = $Matches[1].Split(";")
                            $Width = 100 - [int]$Bar[0]
                            If (-not $Matches[2])
                            {   $Text = "&nbsp;"
                            }
                            Else
                            {   $Text = $Matches[2]
                            }
                            $Line = $Line.Replace($Matches[0],"<td><div style=""background-color:$($Bar[1]);float:left;width:$($Bar[0])%"">$Text</div><div style=""background-color:$($Bar[2]);float:left;width:$width%"">&nbsp;</div></td>")
                            Break
                        }
                        "\[image:(.*?)\](.*?)<\/td>"
                        {   $Image = $Matches[1].Split(";")
                            $Line = $Line.Replace($Matches[0],"<img src=""$($Image[2])"" alt=""$($Matches[2])"" height=""$($Image[0])"" width=""$($Image[1])""></td>")
                        }
                        "\[link:(.*?)\](.*?)<\/td>"
                        {   $Line = $Line.Replace($Matches[0],"<a href=""$($Matches[1])"">$($Matches[2])</a></td>")
                        }
                        "\[linkpic:(.*?)\](.*?)<\/td>"
                        {   $Images = $Matches[1].Split(";")
                            $Line = $Line.Replace($Matches[0],"<a href=""$($Matches[2])""><img src=""$($Image[2])"" height=""$($Image[0])"" width=""$($Image[1])""></a></td>")
                        }
                        Default
                        {   Break
                        }
                    }
                } Until ($Line -notmatch "\[.*?\]")
                $NewHTML += $Line
            }
            Return $NewHTML
        }
    }
#endregion Functions

#region Execution
    Import-Module  DataProtectionManager

    #region bodyTop
        $bodyTop = @"
	        <body>
		        <table style="width: $($tableWidth)px;" width="$tableWidth">
			        <tr>
				        <td style="width: 80%;border: none;" class="headerTable">Azure Backup Report</td>
			        </tr>
		        </table>
"@
    #endregion bodyTop
    
    #region SubHead
        #region Table Begin
            $subHead01 = @"
		            <table cellspacing="0" cellpadding="0" class="inner" border="0" style="margin: 0px; width: $tableWidth;">
			            <tr>
				            <td class="subheader" style="border-top: none;border-bottom: none;">
"@ 
        #endregion Table Begin

        #region Table End
            $subHead02 = @"
				            </td>
			            </tr>
		            </table>
"@
        #endregion Table End
    #endregion SubHead

    Connect-DPMServer -DPMServerName $DPMServerName


    #region RecoveryPoints
        $ProdServers = Get-DPMProductionServer 
    
        $DPMAlerts = Get-DPMAlert -IncludeAlerts AllActive
        $Report = ''

        $listRecoveryPoints = @()
        $i = 0
        $RecoveryPointErrors = 0

        $DPMProtectionGroups = Get-DPMProtectionGroup
        foreach ($DPMProtectionGroup in $DPMProtectionGroups)
        {
            $i++
            Write-Progress -Activity "loop throught Protection Groups" -Status "Server $($DPMProtectionGroup.Name) $i / $($DPMProtectionGroups.Count)" -PercentComplete ( $i / $DPMProtectionGroups.Count * 100) -Id 0

            $DataSources = $DPMProtectionGroup.GetDatasources()
            # $DataSources | Out-GridView
            Write-Verbose "Loop through each available DataSource"
            $j = 0
            foreach ($DataSource in $DataSources)
            {
                $j++
                Write-Progress -Activity "loop throught Data Sources" -Status "DataSource $($DataSource.Name) $j / $($DataSources.Count)" -PercentComplete ( $j / $DataSources.Count * 100) -Id 1 -ParentId 0

                Write-Verbose "Get a list of RecoveryPoints for each DataSource"
                $RecoveryPoints = Get-RecoveryPoint -Datasource $DataSource
                $RecoveryPointsOnline = Get-RecoveryPoint -Datasource $DataSource -Online
                $DPMAlert = ''
                $DPMAlert = ( $DPMAlerts | where { $_.Server -eq $DataSource.ProductionServerName -and $_.AffectedArea -eq $DataSource.Name } ).ErrorInfo.ShortProblem -join ', '

                
                if ( $DPMAlert -ne '' )
                {
                    $DPMAlert = "[cell:$colorRed]$DPMAlert"
                    $RecoveryPointErrors++
                }
                if ( @($RecoveryPoints).Count -ne 0 )
                {
                    $LastPoint = Get-Date ( ( $RecoveryPoints | Sort-Object RepresentedPointInTime | Select -Last 1 ).RepresentedPointInTime ) -Format "dd.MM.yyyy HH:mm:ss"
                    if ( ( $RecoveryPoints | Sort-Object RepresentedPointInTime | Select -Last 1 ).RepresentedPointInTime -gt ( Get-Date ).AddDays( 0 - $RecoveryPointAgeOK ) )
                    {
                        $LastPoint = "[cell:$colorGreen]$LastPoint"
                    } elseif ( ( $RecoveryPoints | Sort-Object RepresentedPointInTime | Select -Last 1 ).RepresentedPointInTime -gt ( Get-Date ).AddDays( 0 - $RecoveryPointAgeWarning ) ) {
                        $LastPoint = "[cell:$colorYellow]$LastPoint"
                    } else {
                        $LastPoint = "[cell:$colorRed]$LastPoint"
                    }
                } else {
                    $LastPoint = "[cell:$colorRed]never"
                }
                if ( $DataSource.IsPresentOnCloud )
                {
                    if ( @($RecoveryPointsOnline).Count -ne 0 )
                    {
                        $LastPointOnline = Get-Date ( ( $RecoveryPointsOnline | Sort-Object RepresentedPointInTime | Select -Last 1 ).RepresentedPointInTime ) -Format "dd.MM.yyyy HH:mm:ss"
                        if ( ( $RecoveryPointsOnline | Sort-Object RepresentedPointInTime | Select -Last 1 ).RepresentedPointInTime -gt ( Get-Date ).AddDays( 0 - $RecoveryPointAgeOK ) )
                        {
                            $LastPointOnline = "[cell:$colorGreen]$LastPointOnline"
                        } elseif ( ( $RecoveryPointsOnline | Sort-Object RepresentedPointInTime | Select -Last 1 ).RepresentedPointInTime -gt ( Get-Date ).AddDays( 0 - $RecoveryPointAgeWarning ) ) {
                            $LastPointOnline = "[cell:$colorYellow]$LastPointOnline"
                        } else {
                            $LastPointOnline = "[cell:$colorRed]$LastPointOnline"
                        }

                    } else {
                        $LastPointOnline = "[cell:$colorRed]never"
                    }
                } else {
                    $LastPointOnline = "not configured"
                }

                $LineItem = New-Object -TypeName PSobject -Property @{
                    "Recovery Points"  = @( $RecoveryPoints ).Count
                    "First Point"      = ( $RecoveryPoints | Sort-Object RepresentedPointInTime | Select -First 1 ).RepresentedPointInTime
                    "Last Point"       = $LastPoint
                    "Size GB"          = [int]( ( $RecoveryPoints | Sort-Object RepresentedPointInTime | Select -Last 1 ).Size / 1GB * 100 ) / 100
                    "Cloud Recovery Points"  = @( $RecoveryPointsOnline ).Count
                    "Cloud First Point"      = ( $RecoveryPointsOnline | Sort-Object RepresentedPointInTime | Select -First 1 ).RepresentedPointInTime
                    "Cloud Last Point"       = $LastPointOnline
                    "Cloud Size GB"          = $CloudSize
                    Server             = "[cellclass:overflowhidden]$($DataSource.ProductionServerName)"
                    Name               = "[cellclass:overflowhidden]$($DataSource.Name -replace "_", "_ ")"
                    "Protection Group" = $DataSource.ProtectionGroup.Name
                    "Object Type"      = $DataSource.ObjectType
                    Alerts             = $DPMAlert
                    Activity           = $DataSource.Activity
                    State              = $DataSource.State
                }
                $listRecoveryPoints += $LineItem
            }
        }
        $ReportRecoveryPoints = $listRecoveryPoints | Sort-Object Server, Name | Select Server, Name, "Object Type", "Protection Group", "Recovery Points", "First Point", "Last Point", "Size GB", "Cloud Recovery Points", "Cloud First Point", "Cloud Last Point", Alerts | ConvertTo-AdvHTML -HeadWidth 72,146,72,72,72,72,72,52,52,72,72,137 -Fragment
        $ReportRecoveryPoints = $subHead01 + "Recovery Points" + $subHead02 + $ReportRecoveryPoints

        Write-Progress -Activity "loop throught Data Sources" -Status "completed" -id 1 -Completed
        Write-Progress -Activity "loop throught Protection Groups" -Status "completed" -id 0 -Completed
    #endregion RecoveryPoints

    #region AgentStatus
        if ( $showAgentStatus -eq $true )
        {
            $AgentStatus = @()
            $i = 0
            foreach ($ProdServer in $ProdServers)
            {
                $i++
                Write-Progress -Activity "Update Agent Status" -Status "Agent $($ProdServer.ServerName) $i / $($ProdServers.Count)" -PercentComplete ( $i / $ProdServers.Count * 100)
                Update-DPMProductionServer -ProductionServer $ProdServer -ErrorAction SilentlyContinue
            }
            $ProdServers = Get-DPMProductionServer
            $i = 0
            foreach ($ProdServer in $ProdServers)
            {
                $i++
                Write-Progress -Activity "Update Agent Status" -Status "Agent $($ProdServer.ServerName) $i / $($ProdServers.Count)" -PercentComplete ( $i / $ProdServers.Count * 100)
                
                if ( $ProdServer.ServerProtectionState -ne 'Deleted' -or $showAgentsDeleted -eq $true ) 
                {
                    switch ($ProdServer.Connectivity.Status)
                    {
                        'OK'              { $ServerConnectivity = "[cell:$colorGreen]OK" }
                        'Error'           { $ServerConnectivity = "[cell:$colorRed]Error" }
                        'Unavailable'     { $ServerConnectivity = "[cell:$colorRed]Unavailable" }
                        'Restart pending' { $ServerConnectivity = "[cell:$colorYellow]Restart pending" }
                        Default           { $ServerConnectivity = "[cell:$colorYellow]$($ProdServer.Connectivity)" }
                                                    
                    }
                    if ($ProdServer.UpgradeAvailable -eq $true)
                    {
                        $UpgradeAvailable = "[cell:$colorYellow]Yes"
                    } else {
                        $UpgradeAvailable = "[cell:$colorGreen]No"
                    }

                    $IsThrottled = $ProdServer.IsThrottled
                    switch ($ProdServer.ServerProtectionState)
                    {
                        'HasDatasourcesProtected' {
                            $ServerProtectionState = "[cell:$colorGreen]has Datasources protected"
                        }
                        'Deleted'                { 
                            $ServerProtectionState = "[cell:$colorGray]deleted" 
                            $ServerConnectivity    = "[cell:$colorGray]"
                            $UpgradeAvailable      = "[cell:$colorGray]"
                            $IsThrottled           = "[cell:$colorGray]"
                        }
                        'NoDatasourcesProtected' { $ServerProtectionState = "[cell:$colorRed]no Datasources protected" }

                    }
                    $AgentStatus += New-Object -TypeName PSObject -Property @{
                        "Server Name"       = $ProdServer.MachineName
                        "Protection State"  = $ServerProtectionState
                        "Connectivity"      = $ServerConnectivity
                        "Upgrade Available" = $UpgradeAvailable
                        IsThrottled         = $IsThrottled

                    }
                }
            }
        
            $ReportAgentStatus = $AgentStatus | select "Server Name", "Protection State", Connectivity, "Upgrade Available", IsThrottled | Sort-Object "Server Name" | ConvertTo-AdvHTML -HeadWidth 197,197,197,197,196 -Fragment
            $ReportAgentStatus = $subHead01 + "Agents" + $subHead02 + $ReportAgentStatus
        } else {
            $ReportAgentStatus = ''
        }
        Write-Progress -Activity "Update Agent Status" -Status "completed" -id 0 -Completed
    #endregion AgentStatus

    #region DiskStatus
        if ( $showDisks -eq $true )
        {
            $DPMDisks = Get-DPMDisk | where { $_.IsInStoragePool -eq $true }
            $ReportDisks = @()
            foreach ($DPMDisk in $DPMDisks)
            {
                $percent = @( ( ( $DPMDisk.TotalCapacity - $DPMDisk.UnallocatedSpace ) / $DPMDisk.TotalCapacity * 100 ) -split "\." )[0]
                if ( $percent -le 70 )
                {
                    $barcolor = $colorGreen
                } elseif ( $percent -le 90 ) {
                    $barcolor = $colorYellow
                } else {
                    $barcolor = $colorRed
                }
                $percentUsed = "[bar:$( $percent + 1 );$barcolor;#FFFFFF]$( $percent.ToString() ) % used"
                $ReportDisks += New-Object -TypeName PSObject -Property @{
                    Name = "Disk $($DPMDisk.NtDiskId) ($($DPMDisk.Name))"
                    Capacity = $DPMDisk.TotalCapacityLabel
                    'Free Space' = $DPMDisk.UnallocatedSpaceLabel
                    'Disk Status' = $DPMDisk.DiskStatus
                    'Percent Used' = $percentUsed

                }
            }


            $ReportDisks += New-Object -TypeName PSObject -Property @{
                    Name = "Total"
                    Capacity = ( ( ( $DPMDisks | Measure-Object TotalCapacity -Sum ).Sum / 1GB ) -split "\." )[0] + " GB"
                    'Free Space' = ( ( ( $DPMDisks | Measure-Object UnallocatedSpace -Sum ).Sum / 1GB ) -split "\." )[0] + " GB"
                    'Disk Status' = ''
                    'Percent Used' = ''

                }            

            # $ReportDisks = $ReportDisks | Select Name, Capacity, 'Free Space', 'Percent Used', 'Disk Status' | ConvertTo-AdvHTML -HeadWidth 197,197,197,197,196 -Fragment
            $ReportDisks = $ReportDisks | Select Name, Capacity, 'Free Space', 'Disk Status' | ConvertTo-AdvHTML -HeadWidth 397,197,197,196 -Fragment
            $ReportDisks = $subHead01 + "Disks" + $subHead02 + $ReportDisks
        } else {
            $ReportDisks = ''
        }
    #endregion DiskStatus

    #region Jobs in Progress
        if ( $showJobsInProgress -eq $true )
        {
            $DPMJobsInProgress = Get-DPMJob -Status InProgress -DPMServerName $DPMServerName 
            # $DPMJobsInProgress.Count
            $DPMJobsTasks = @()
            foreach ( $DPMJobInProgress in $DPMJobsInProgress ) 
            {
                foreach ( $DPMJobTask in $DPMJobInProgress.TaskList )
                {
                    # $DPMJobsTasks += $DPMJobTask
                    if ( $DPMJobTask.StartTime -eq (Get-Date "01.01.0001") )
                    {
                        $timeElapsed = New-TimeSpan -Start $DPMJobTask.CreatedTime -End ( Get-Date )
                        $timeElapsed = @( ( "{0:G}" -f $timeElapsed ) -split "\." )[0]
                        $StartTime   = ''
                    } else {
                        $timeElapsed = New-TimeSpan -Start $DPMJobTask.StartTime -End ( Get-Date )
                        $timeElapsed = @( ( "{0:G}" -f $timeElapsed ) -split "\." )[0]
                        $StartTime   = $DPMJobTask.StartTime
                    }
                    $DPMJobsTasks += New-Object -TypeName PSObject -Property @{
                        Created  = $DPMJobTask.CreatedTime
                        'Start Time' = $StartTime
                        'Time Elapsed' = $timeElapsed
                        Status = $DPMJobTask.Status
                        Type   = $DPMJobTask.Type
                        Server = $DPMJobTask.ProductionServerName
                        Transfered = "$( [int]($DPMJobInProgress.DataSize / 1GB * 100) / 100 ) GB"
                        Path = "[cellclass:overflowhidden]$($DPMJobTask.DatasourcePath)"
                        Error = $DPMJobTask.ErrorInfo.ShortProblem
                    }
                }
            }
            $ReportJobsTasks = $DPMJobsTasks | Select Created, 'Start Time', 'Time Elapsed', Status, Transfered, Type, Server, Path | ConvertTo-AdvHTML -HeadWidth 117,117,117,187,97,97,147,96 -Fragment
            $ReportJobsTasks = $subHead01 + "Jobs in Progress" + $subHead02 + $ReportJobsTasks 
        } else {
            $ReportJobsTasks = ''
        }
    #endregion Jobs in Progress

    #region Jobs Failed
        if ( $showJobsFailed -eq $true )
        {
            $DPMJobsFailed = Get-DPMJob -Status Failed -From ( Get-Date ( ( Get-Date ).AddDays( -1 ) ) -Format "dd.MM.yyyy" )
            $DPMJobsFailedTasks = @()
            foreach ( $DPMJobFailed in $DPMJobsFailed ) 
            {
                # $DPMJobFailed = $DPMJobsFailed[0]
                foreach ( $DPMJobTask in $DPMJobFailed.TaskList )
                {
                    # $DPMJobsFailedTasks += $DPMJobTask
                    $DPMJobsFailedTasks += New-Object -TypeName PSObject -Property @{
                        Created  = $DPMJobTask.CreatedTime
                        'Start Time' = $DPMJobTask.StartTime
                        'End Time' = $DPMJobTask.EndTime
                        Status = $DPMJobTask.Status
                        Type   = $DPMJobTask.Type
                        Server = $DPMJobTask.ProductionServerName
                        Path = "[cellclass:overflowhidden]$($DPMJobTask.DatasourcePath)"
                        Error = $DPMJobTask.ErrorInfo.ShortProblem
                    }
                }
            }
            $ReportJobsTasksFailed = $DPMJobsFailedTasks | Select Created, 'Start Time', 'End Time', Status, Type, Server, Path, Error | ConvertTo-AdvHTML -HeadWidth 117,117,117,77,97,147,97,206 -Fragment
            $ReportJobsTasksFailed = $subHead01 + "Failed Jobs for yesterday and today" + $subHead02 + $ReportJobsTasksFailed 
        } else {
            $ReportJobsTasksFailed = ''
        }
    #endregion Jobs Failed

    #region Azure
       <# Import-Module MSOnlineBackup
        $OBMachineUsage = [int]( ( Get-OBMachineUsage ).StorageUsedByMachineInBytes / 1GB * 100) / 100

        $ReportAzure = @"
            <table style="width: $tableWidth;">
                <tr>
                    <td>
                        Total Space used in Azure:
                    </td>
                    <td>
                        $OBMachineUsage GB
                    </td>
                </tr>
            </table>
"@
        $ReportAzure = $subHead01 + "Azure Backup" + $subHead02 + $ReportAzure 
#>
    #endregion Azure

    #region Overall-State
        if ( 
            ( $RecoveryPointErrors -ne 0 ) -or `
            ( @( $ProdServers | where { $_.Connectivity -eq 'Error' } ).Count -ne 0 ) -or  `
            ( @( $ProdServers | where { $_.Connectivity -eq 'Unavailable' } ).Count -ne 0 )
        ) {
            $StateColor = $colorRed
            $OverAllState = 'Failed'
        } elseif ( 
            (@( $AgentStatus | where { $_."Upgrade Available" -match 'Yes' } ).Count -ne 0 ) -or  `
            (@( $ProdServers | where { $_.Connectivity.Status -eq 'RebootRequired' } ).Count -ne 0 ) -or `
            (@( $ProdServers | where { $_.Connectivity.Status -eq 'Unknown' } ).Count -ne 0 )
        ) {
            $StateColor = $colorYellow
            $OverAllState = 'Warning'
        } else {
            $StateColor = $colorGreen
            $OverAllState = 'Success'
        }
    #endregion Overall-State
    
    #region Header
        $headerObj = @"
        <html>
	        <head>
		        <title>$emailSubject</title>
		        <style>  
			        body {font-family: Tahoma; background-color:#fff;width: 1000px;}
			        h1.top {background-color: #fb9895;color: White;font-weight: bold;font-size: 16px;vertical-align: center;padding: 5px;}
			        table {font-family: Tahoma;font-size: 12px;background-color: #e3e3e3;width:$($tableWidth)px;border-collapse:collapse; table-layout: fixed;}
			        .headerTable{background-color: $StateColor ;color: White;font-weight: bold;font-size: 16px;height: 70px;vertical-align: bottom;padding: 0 0 15px 15px;border-bottom: none;}
			        .subheader{height: 35px;background-color: #f3f4f4;font-size: 16px;vertical-align: middle;padding: 5px 0 0 15px;color: #626365;}
			        table.inner tr{height: 17px;}
			        table.inner td{padding: 2px 2px 2px 2px;vertical-align: top;border: 1px solid #a7a9ac; }
			        table.inner .subheader{height: 35px;background-color: #f3f4f4;font-size: 16px;vertical-align: middle;padding: 5px 0 0 15px;color: #626365;}
                    td.overflowhidden { overflow: hidden; }
			        th {border: 1px solid #a7a9ac;border-bottom: none;}
			        tr {height: 17px;}
			        td {background-color: #fff;border: 1px solid #a7a9ac;padding: 2px 2px 2px 2px;vertical-align: top;}  
                </style>
            </head>
"@
    #endregion Header

    #region Footer
        $bodyScript = @"
            <table style="width: $tableWidth;">
            <tr><td>This report was generated by Script <b>$( $MyInvocation.InvocationName )</b> on Server <b>$env:computername</b></td></tr>
            </table>
"@
    #endregion Footer

    #region composite Output
        $htmlOutput = $headerObj + `
                      $bodyTop + `
                      $ReportRecoveryPoints + `
                      $ReportJobsTasks + `
                      $ReportJobsTasksFailed + `
                      $ReportAgentStatus + `
                      $ReportDisks + `
#                      $ReportAzure + `
                      $footerObj + `
                      $bodyScript
    #endregion composite Output

    #region Output
        if ($sendEmail -eq $true) 
        {
            $emailSubject = "[$OverAllState] $emailSubject"

            Send-MailMessage -From $emailFrom -To $emailTo -Subject $emailSubject -Body $htmlOutput -BodyAsHTML -SmtpServer $emailHost
        } 
        if ( $saveHTML -eq $true )
        {
	        $htmlOutput | Out-File $file
            # [Reflection.Assembly]::LoadWithPartialName("System.Xml.Linq")
            # [System.Xml.Linq.XDocument]::Load($file).Save($file)
	        Invoke-Expression $file
        }
    #endregion
#endregion Execution

#region ScriptEnd
    Write-Host "`r`n Ended at $(get-date)"
    Write-Host "`r`n Total Elapsed Time: $($elapsed.Elapsed.ToString())"
#endregion ScriptEnd