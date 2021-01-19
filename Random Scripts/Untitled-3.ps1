$groups = Get-ADGroup -Filter * -Properties member, Description -SearchBase "OU=Grant Local Admin,OU=Groups,OU=DDIA,DC=deltadentalia,DC=com"

$array = foreach ($group in $groups)
{
    $users = Get-ADGroupMember $group -Recursive
 	
    foreach($user in $users)
    {
        $user = Get-aduser $user -Properties samaccountname
        @{
            Groups = $group.Name
            Description = $group.Description
            Users = $user.samaccountname
        } 
    }
}
$array | ForEach-Object { [pscustomobject] $_ } | sort Groups | Export-Csv C:\TEMP\groups_with_users.csv -NoTypeInformation