$computers= get-content c:\temp\computerlist.txt  

$bdeObject = @()
foreach ($computer in $computers) {
        $bde = manage-bde -cn $computer -status
            $ComputerName = $bde | Select-String "Computer Name:" 
            $ComputerName = ($ComputerName -split ": ")[1]

            $ConversionStatus = $bde | Select-String "Conversion Status:"
            $ConversionStatus = ($ConversionStatus -split ": ")[1]
            $ConversionStatus = $ConversionStatus -replace '\s','' #removes the white space in this field

            $PercentageEncrypted = $bde | Select-String "Percentage Encrypted:"
            $PercentageEncrypted = ($PercentageEncrypted -split ": ")[1]

        #Add all fields to an array that contains custom formatted objects with desired fields
        $bdeObject += New-Object psobject -Property @{'Computer Name'=$ComputerName; 'Conversion Status'=$ConversionStatus; 'Percentage Encrypted'=$PercentageEncrypted;}
    }

$bdeObject | sort "computer name"