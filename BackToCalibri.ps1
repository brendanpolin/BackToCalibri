# Search for directories named 'BackToCalibri' that are under any 'DropZone' directory
$rootPath = "C:\"
$searchPattern = "*\BackToCalibri"

# Perform the search
$foundDirectories = Get-ChildItem -Path $rootPath -Recurse -ErrorAction SilentlyContinue -Directory | Where-Object { $_.FullName -like "$rootPath\$searchPattern" }

#Sets a variable to the found directory that has the share
$calibriDirPath = $foundDirectories.FullName

#Creates variables with the paths for ease of use
$PowerpointPath = Join-Path -Path $calibriDirPath -ChildPath "\Blank.potx"
$WordPath = Join-Path -Path $calibriDirPath -ChildPath "\Normal.dotm"

#BEGIN Excel Registry for all users
#Script from https://www.pdq.com/blog/modifying-the-registry-users-powershell/

# Regex pattern for SIDs
#$PatternSID = 'S-1-5-21-\d+-\d+\-\d+\-\d+$'
$PatternSID = @('S-1-5-21-\d+-\d+\-\d+\-\d+$',	 # SID of AD Accounts
				            'S-1-12-1-\d+-\d+-\d+-\d+$')	   # SID of Entra Accounts
 
# Get Username, SID, and location of ntuser.dat for all users
Foreach ($pattern in $PatternSID) {
    $ProfileList += Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object {$_.PSChildName -match $pattern} | 
        Select  @{name="SID";expression={$_.PSChildName}}, 
                @{name="UserHive";expression={"$($_.ProfileImagePath)\ntuser.dat"}}, 
                @{name="Username";expression={$_.ProfileImagePath -replace '^(.*[\\\/])', ''}}
    # Get all user SIDs found in HKEY_USERS (ntuder.dat files that are loaded)
    $LoadedHives += Get-ChildItem Registry::HKEY_USERS | ? {$_.PSChildname -match $pattern} | Select-Object @{name="SID";expression={$_.PSChildName}}
}

#This checks their SID length to make sure it's not a service account, and if it's not then it will copy over the templates
Foreach ($item in $ProfileList) {
    if ($item.SID.Length -gt 10) {
        $Username = $item.Username
        Copy-Item $PowerpointPath C:\Users\$Username\AppData\Roaming\Microsoft\Templates
        Copy-Item $WordPath C:\Users\$Username\AppData\Roaming\Microsoft\Templates
    }
}
 
# Get all users that are not currently logged
$UnloadedHives = Compare-Object $ProfileList.SID $LoadedHives.SID | Select-Object @{name="SID";expression={$_.InputObject}}, UserHive, Username
 
#-----This is all for outlook Calibri --------------------
$ValueSimple = "3C,00,00,00,1F,00,00,F8,00,00,00,40,DC,00,00,00,00,00,00,00,00,00,00,FF,00,22,43,61,6C,69,62,72,69,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
$ValueComposeComplex = "3C,68,74,6D,6C,3E,0D,0A,0D,0A,3C,68,65,61,64,3E,0D,0A,3C,73,74,79,6C,65,3E,0D,0A,0D,0A,20,2F,2A,20,53,74,79,6C,65,20,44,65,66,69,6E,69,74,69,6F,6E,73,20,2A,2F,0D,0A,20,73,70,61,6E,2E,50,65,72,73,6F,6E,61,6C,43,6F,6D,70,6F,73,65,53,74,79,6C,65,0D,0A,09,7B,6D,73,6F,2D,73,74,79,6C,65,2D,6E,61,6D,65,3A,22,50,65,72,73,6F,6E,61,6C,20,43,6F,6D,70,6F,73,65,20,53,74,79,6C,65,22,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,74,79,70,65,3A,70,65,72,73,6F,6E,61,6C,2D,63,6F,6D,70,6F,73,65,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,6E,6F,73,68,6F,77,3A,79,65,73,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,75,6E,68,69,64,65,3A,6E,6F,3B,0D,0A,09,6D,73,6F,2D,61,6E,73,69,2D,66,6F,6E,74,2D,73,69,7A,65,3A,31,31,2E,30,70,74,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,73,69,7A,65,3A,31,31,2E,30,70,74,3B,0D,0A,09,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,43,61,6C,69,62,72,69,22,2C,73,61,6E,73,2D,73,65,72,69,66,3B,0D,0A,09,6D,73,6F,2D,61,73,63,69,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,43,61,6C,69,62,72,69,3B,0D,0A,09,6D,73,6F,2D,68,61,6E,73,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,43,61,6C,69,62,72,69,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,54,69,6D,65,73,20,4E,65,77,20,52,6F,6D,61,6E,22,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,74,68,65,6D,65,2D,66,6F,6E,74,3A,6D,69,6E,6F,72,2D,62,69,64,69,3B,0D,0A,09,63,6F,6C,6F,72,3A,77,69,6E,64,6F,77,74,65,78,74,3B,0D,0A,09,66,6F,6E,74,2D,77,65,69,67,68,74,3A,6E,6F,72,6D,61,6C,3B,0D,0A,09,66,6F,6E,74,2D,73,74,79,6C,65,3A,6E,6F,72,6D,61,6C,3B,7D,0D,0A,2D,2D,3E,0D,0A,3C,2F,73,74,79,6C,65,3E,0D,0A,3C,2F,68,65,61,64,3E,0D,0A,0D,0A,3C,2F,68,74,6D,6C,3E,0D,0A"
$ValueReplyComplex = "3C,68,74,6D,6C,3E,0D,0A,0D,0A,3C,68,65,61,64,3E,0D,0A,3C,73,74,79,6C,65,3E,0D,0A,0D,0A,20,2F,2A,20,53,74,79,6C,65,20,44,65,66,69,6E,69,74,69,6F,6E,73,20,2A,2F,0D,0A,20,73,70,61,6E,2E,50,65,72,73,6F,6E,61,6C,52,65,70,6C,79,53,74,79,6C,65,0D,0A,09,7B,6D,73,6F,2D,73,74,79,6C,65,2D,6E,61,6D,65,3A,22,50,65,72,73,6F,6E,61,6C,20,52,65,70,6C,79,20,53,74,79,6C,65,22,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,74,79,70,65,3A,70,65,72,73,6F,6E,61,6C,2D,72,65,70,6C,79,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,6E,6F,73,68,6F,77,3A,79,65,73,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,75,6E,68,69,64,65,3A,6E,6F,3B,0D,0A,09,6D,73,6F,2D,61,6E,73,69,2D,66,6F,6E,74,2D,73,69,7A,65,3A,31,31,2E,30,70,74,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,73,69,7A,65,3A,31,31,2E,30,70,74,3B,0D,0A,09,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,43,61,6C,69,62,72,69,22,2C,73,61,6E,73,2D,73,65,72,69,66,3B,0D,0A,09,6D,73,6F,2D,61,73,63,69,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,43,61,6C,69,62,72,69,3B,0D,0A,09,6D,73,6F,2D,68,61,6E,73,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,43,61,6C,69,62,72,69,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,54,69,6D,65,73,20,4E,65,77,20,52,6F,6D,61,6E,22,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,74,68,65,6D,65,2D,66,6F,6E,74,3A,6D,69,6E,6F,72,2D,62,69,64,69,3B,0D,0A,09,63,6F,6C,6F,72,3A,77,69,6E,64,6F,77,74,65,78,74,3B,0D,0A,09,66,6F,6E,74,2D,77,65,69,67,68,74,3A,6E,6F,72,6D,61,6C,3B,0D,0A,09,66,6F,6E,74,2D,73,74,79,6C,65,3A,6E,6F,72,6D,61,6C,3B,7D,0D,0A,2D,2D,3E,0D,0A,3C,2F,73,74,79,6C,65,3E,0D,0A,3C,2F,68,65,61,64,3E,0D,0A,0D,0A,3C,2F,68,74,6D,6C,3E,0D,0A"
$ValueTextComplex = "3C,68,74,6D,6C,3E,0D,0A,0D,0A,3C,68,65,61,64,3E,0D,0A,3C,73,74,79,6C,65,3E,0D,0A,0D,0A,20,2F,2A,20,53,74,79,6C,65,20,44,65,66,69,6E,69,74,69,6F,6E,73,20,2A,2F,0D,0A,20,70,2E,4D,73,6F,50,6C,61,69,6E,54,65,78,74,2C,20,6C,69,2E,4D,73,6F,50,6C,61,69,6E,54,65,78,74,2C,20,64,69,76,2E,4D,73,6F,50,6C,61,69,6E,54,65,78,74,0D,0A,09,7B,6D,73,6F,2D,73,74,79,6C,65,2D,6E,6F,73,68,6F,77,3A,79,65,73,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,70,72,69,6F,72,69,74,79,3A,39,39,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,6C,69,6E,6B,3A,22,50,6C,61,69,6E,20,54,65,78,74,20,43,68,61,72,22,3B,0D,0A,09,6D,61,72,67,69,6E,3A,30,69,6E,3B,0D,0A,09,6D,73,6F,2D,70,61,67,69,6E,61,74,69,6F,6E,3A,77,69,64,6F,77,2D,6F,72,70,68,61,6E,3B,0D,0A,09,66,6F,6E,74,2D,73,69,7A,65,3A,31,31,2E,30,70,74,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,73,69,7A,65,3A,31,30,2E,35,70,74,3B,0D,0A,09,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,43,61,6C,69,62,72,69,22,2C,73,61,6E,73,2D,73,65,72,69,66,3B,0D,0A,09,6D,73,6F,2D,66,61,72,65,61,73,74,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,43,61,6C,69,62,72,69,3B,0D,0A,09,6D,73,6F,2D,66,61,72,65,61,73,74,2D,74,68,65,6D,65,2D,66,6F,6E,74,3A,6D,69,6E,6F,72,2D,6C,61,74,69,6E,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,54,69,6D,65,73,20,4E,65,77,20,52,6F,6D,61,6E,22,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,74,68,65,6D,65,2D,66,6F,6E,74,3A,6D,69,6E,6F,72,2D,62,69,64,69,3B,7D,0D,0A,2D,2D,3E,0D,0A,3C,2F,73,74,79,6C,65,3E,0D,0A,3C,2F,68,65,61,64,3E,0D,0A,0D,0A,3C,2F,68,74,6D,6C,3E,0D,0A"
 
$registryPath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\mailsettings'
$Name1Simple = "ComposeFontSimple"
$Name1Complex = "ComposeFontComplex"
$Name2Simple = "ReplyFontSimple"
$Name2Complex = "ReplyFontComplex"
$Name3Simple = "TextFontSimple"
$Name3Complex = "TextFontComplex"
 
$hexSimple = $ValueSimple.Split(',') | % { "0x$_"}
$hexComposeComplex = $ValueComposeComplex.Split(',') | % { "0x$_"}
$hexReplyComplex = $ValueReplyComplex.Split(',') | % { "0x$_"}
$hexTextComplex = $ValueTextComplex.Split(',') | % { "0x$_"}
 #-----This is all for outlook Calibri --------------------

# Loop through each profile on the machine
Foreach ($item in $ProfileList) {
    # Load User ntuser.dat if it's not already loaded
    IF ($item.SID -in $UnloadedHives.SID) {
        reg load HKU\$($Item.SID) $($Item.UserHive) | Out-Null
    }
 
    #####################################################################
    # This is where you can read/modify a users portion of the registry 
 
    # Changes the default font to Calibri, 11 in Excel
    "{0}" -f $($item.Username) | Write-Output
    Set-ItemProperty registry::HKEY_USERS\$($Item.SID)\Software\Microsoft\Office\16.0\Excel\Options -Name font -Value "Calibri, 11" -Type String -Force
    $registryPath = "registry::HKEY_USERS\$($Item.SID)\SOFTWARE\Microsoft\Office\16.0\Common\mailsettings"

    IF(!(Test-Path $registryPath))
    {
    New-Item -Path $registryPath -Force | Out-Null
    New-ItemProperty -Path $registryPath -name NewTheme -PropertyType string
    New-ItemProperty -Path $registryPath -Name $Name1Simple -Value ([byte[]]$hexSimple) -PropertyType Binary -Force
    New-ItemProperty -Path $registryPath -Name $Name2Simple -Value ([byte[]]$hexSimple) -PropertyType Binary -Force
    New-ItemProperty -Path $registryPath -Name $Name3Simple -Value ([byte[]]$hexSimple) -PropertyType Binary -Force
    New-ItemProperty -Path $registryPath -Name $Name1Complex -Value ([byte[]]$hexComposeComplex) -PropertyType Binary -Force
    New-ItemProperty -Path $registryPath -Name $Name2Complex -Value ([byte[]]$hexReplyComplex) -PropertyType Binary -Force
    New-ItemProperty -Path $registryPath -Name $Name3Complex -Value ([byte[]]$hexTextComplex) -PropertyType Binary -Force
    }
   
    ELSE {
    Set-ItemProperty -Path $registryPath -name NewTheme -value $null
    Set-ItemProperty -Path $registryPath -name ThemeFont -value 2
    Set-ItemProperty -Path $registryPath -Name $Name1Simple -Value ([byte[]]$hexSimple) -Force
    Set-ItemProperty -Path $registryPath -Name $Name2Simple -Value ([byte[]]$hexSimple) -Force
    Set-ItemProperty -Path $registryPath -Name $Name3Simple -Value ([byte[]]$hexSimple) -Force
    Set-ItemProperty -Path $registryPath -Name $Name1Complex -Value ([byte[]]$hexComposeComplex) -Force
    Set-ItemProperty -Path $registryPath -Name $Name2Complex -Value ([byte[]]$hexReplyComplex) -Force
    Set-ItemProperty -Path $registryPath -Name $Name3Complex -Value ([byte[]]$hexTextComplex) -Force
    }
    
    #####################################################################
 
    # Unload ntuser.dat        
    IF ($item.SID -in $UnloadedHives.SID) {
        ### Garbage collection and closing of ntuser.dat ###
        [gc]::Collect()
        reg unload HKU\$($Item.SID) | Out-Null
    }
}
