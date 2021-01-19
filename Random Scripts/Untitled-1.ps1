Get-ADUser -Filter * -SearchBase "OU=Users,OU=DDIA,DC=deltadentalia,DC=com" -Properties Name, manager | Where-Object {($_.manager -eq $Null) -and ($_.Enabled -eq 'True') -and ($_.DistinguishedName -notlike "*,OU=Process Accounts,OU=Users,OU=DDIA,DC=deltadentalia,DC=com")} | select name