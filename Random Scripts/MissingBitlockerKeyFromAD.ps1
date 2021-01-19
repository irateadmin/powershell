$havekey = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -searchbase "OU=Windows Defender,OU=Computers,OU=DDIA,DC=deltadentalia,DC=com"| ForEach-Object {$_.distinguishedname.split(',',2)[1]}
$all = get-adcomputer -filter {name -notlike "*covmdt*"} -searchbase "OU=Windows Defender,OU=Computers,OU=DDIA,DC=deltadentalia,DC=com"
  
$uniquekeys = $havekey | select -unique 
  
$nokey = compare-object -ReferenceObject $all.distinguishedname -DifferenceObject $uniquekeys -PassThru | Where-Object {$_.sideindicator -eq "<="}

$array = foreach ($laptop in $nokey) {
            $laptop = get-adcomputer $laptop -properties name,description
            @{
                Name = $laptop.Name
                Description = $laptop.Description
            }
}
    $array | ForEach-Object { [pscustomobject] $_ } | sort