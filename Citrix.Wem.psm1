#
# Citrix.Wem.Version = "0.9.4"
#

<# 
    .Synopsis
    Builds an .xml file containing WEM Action definitions.

    .Description
    Builds an .xml file containing WEM Action definitions for application shortcuts.
    This function supports multiple types of input and creates the file containing the Actions
    ready for import into WEM.

    .Link
    https://msfreaks.wordpress.com

    .Parameter Path
    Can be a targeted folder or a targeted file. If a folder is specified the function
    will assume you wish to parse .lnk files.
    If omitted the default Start Menu\Programs folders (current user and all users)
    locations will be used.

    .Parameter OutputPath
    Location where the output xml file will be written to. Defaults to current folder if
    omitted.

    .Parameter OutputFileName
    The default filename is VUEMApplications.xml. Use this parameter to override this if needed. 

    .Parameter Recurse
    Whether the script needs to recursively process $Path. Only valid when the $Path parameter
    is a folder. Is $True by default if omitted.

    .Parameter FileTypes
    Provide a comma separated list of filetypes to process. Only valid when the $Path
    parameter is a folder. If omitted .lnk will be used by default.

    .Parameter Prefix
    Provide a prefix string used to generate Action names (as displayed in the WEM console).

    .Parameter SelfHealingEnabled
    If used will create Actions using the SelfHealingEnabled parameter. Defaults to $false if omitted.

    .Parameter Disable
    If used will create disabled Actions. Defaults to $false if omitted (create Enabled Actions).

    .Example
    New-VUEMApplicationsXML

    Description

    -----------

    Create WEM Actions from all the items in the default Start Menu locations and export this to VUEMApplications.xml in the current folder.

    .Example
    New-VUEMApplicationsXML -Prefix "ITW - "

    Description

    -----------

    Create WEM Actions from all the items in the default Start Menu locations and export this to VUEMApplications.xml in the current folder.
    All Action names in the WEM Console are prefixed with "ITW - ", like for example "ITW - Notepad".

    .Example
    New-VUEMApplicationsXML -Path "E:\Custom Folder\Menu Items" -FileTypes exe,lnk

    Description

    -----------

    Create VUEMApplications.xml in the current folder from all the items in a custom folder, processing .exe and .lnk files.

    .Example
    New-VUEMApplicationsXML -Path "C:\Windows\System32\notepad.exe" -Name "Notepad example"

    Description

    -----------

    Create WEM Actions from Notepad.exe

    .Example
    New-VUEMApplicationsXML -OutputPath "C:\Temp" -OutputFileName "applications.xml"

    Description

    -----------

    Create applications.xml in c:\temp for all the items in the default Start Menu locations.

    .Example
    New-VUEMApplicationsXML -SelfHealing Enabled -Disable

    Description

    -----------

    Create WEM Actions from all the items in the default Start Menu locations and export this to VUEMApplications.xml in the current folder.
    Actions are created using the Enable SelfHealing switch, and are disabled.

    .Notes
    By default, if no Path is given the default Start Menu locations will be processed.
    If Folder Redirection for the Start Menu folder is detected, that folder will be used instead.
#>
function New-VUEMApplicationsXML {
    param(
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True)][string]$Path,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$OutputFileName = "VUEMApplications.xml",
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][bool]$Recurse = $true,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string[]]$FileTypes,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$Prefix,
        [Parameter(Mandatory=$False,        
        ValueFromPipeline=$False)][switch]$SelfHealingEnabled,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][switch]$Disable
    )

    # check if $Path is valid
    If ($Path -and !(Test-Path $Path)) {
         Write-Host "Cannot find path '$Path' because it does not exist." -ForegroundColor Red
         Break
    }

    # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path $OutputPath) -or ((Get-Item $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    }

    # grab files
    $files = @()
    
    # tell user we're processing the Start Menu if $Path was not provided
    If (!$Path) {
        If ([Environment]::GetFolderPath('StartMenu') -eq "$($env:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Start Menu") {
            Write-Host "`nProcessing default Start Menu folders" -ForegroundColor Yellow
            $files = Get-FilesToProcess("$($env:ProgramData)\Microsoft\Windows\Start Menu\Programs")
            $files += Get-FilesToProcess("$($env:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs")
        } Else {
            Write-Host "`nProcessing redirected Start Menu ($([Environment]::GetFolderPath('StartMenu')))" -ForegroundColor Yellow
            $files = Get-FilesToProcess("$([Environment]::GetFolderPath('StartMenu'))\Programs")
        }
    } Else {
        Write-Host "`nProcessing '$Path'" -ForegroundColor Yellow
        $files = Get-FilesToProcess($Path)
    }

    # pre-load System.Drawing namespace
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    # init VUEM action arrays
    $VUEMApplications = @()

    # define state
    $State = "1"
    If ($Disable) { $State = "0" }

    # define selfhealing
    $SelfHeal = "0"
    If ($SelfHealingEnabled) { $SelfHeal = "1" }

    # process inputs
    ForEach ($file in $files) {
        $Description = ""
        $WorkingDirectory = ""
        $IconLocation = ""
        $Arguments = ""
        $HotKey = "None"
        $TargetPath = ""
        $IconStream = ""
        If ($file.Extension -eq ".lnk") {
            # grab lnk file for property parsing
            $obj = New-Object -ComObject WScript.Shell
            $lnk = $obj.CreateShortcut("$($file.FullName)")
            
            # grab properties
            $Description = $lnk.Description
            $WorkingDirectory = [System.Environment]::ExpandEnvironmentVariables($lnk.WorkingDirectory)
            $IconLocation = [System.Environment]::ExpandEnvironmentVariables($lnk.IconLocation.Split(",")[0])
            $Arguments = $lnk.Arguments
            If ($lnk.Hotkey) { $HotKey = $lnk.Hotkey }
            $TargetPath = [System.Environment]::ExpandEnvironmentVariables($lnk.TargetPath)
            If (!$IconLocation) { $IconLocation = $TargetPath }
        } Else {
            # for anything but .lnk try to generate relevant properties
            $WorkingDirectory = $file.FolderName
            $IconLocation = $file.FullName
            $TargetPath = $file.FullName
        }

        # only work if we have a target
        If ($TargetPath) {
            # grab icon
            If ($IconLocation -and (Test-Path $IconLocation)) { $IconStream = Get-IconToBase64 ([System.Drawing.Icon]::ExtractAssociatedIcon("$($IconLocation)")) }
            
            $VUEMAppName = Get-UniqueActionName -ObjectList $VUEMApplications -ActionName "$Prefix$($file.BaseName)"

            $VUEMApplications += New-VUEMApplicationObject -Name "$VUEMAppName" `
                                                           -Description "$Description" `
                                                           -DisplayName "$($file.BaseName)" `
                                                           -StartMenuTarget "Start Menu\Programs$($file.RelativePath)" `
                                                           -TargetPath "$TargetPath" `
                                                           -Parameters "$Arguments" `
                                                           -WorkingDirectory "$WorkingDirectory" `
                                                           -Hotkey "$Hotkey" `
                                                           -IconLocation "$IconLocation" `
                                                           -IconStream "$IconStream" `
                                                           -SelfHealingEnabled "$SelfHeal" `
                                                           -State "$State"
        }
    }

    # output xml file
    If ($VUEMApplications) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMApplication" -ObjectList $VUEMApplications | Out-File $OutputPath\$OutputFileName
        Write-Host "$OutputFileName written to '$OutputPath\$OutputFileName'" -ForegroundColor Green
    }
}

<# 
    .Synopsis
    Builds an .xml file containing WEM Action definitions.

    .Description
    Builds an .xml file containing WEM Action definitions for Mapped Network Drives for the current user.
    This function creates the file containing the Actions ready for import into WEM.

    .Link
    https://msfreaks.wordpress.com

    .Parameter DriveLetter
    Will only process the Mapped Drive letter defined by "DriveLetter".

    .Parameter OutputPath
    Location where the output xml file will be written to. Defaults to current folder if
    omitted.

    .Parameter OutputFileName
    The default filename is VUEMNetDrives.xml. Use this parameter to override this if needed. 

    .Parameter Prefix
    Provide a prefix string used to generate Action names (as displayed in the WEM console).

    .Parameter SelfHealingEnabled
    If used will create Actions using the SelfHealingEnabled parameter. Defaults to $false if omitted.

    .Parameter Disable
    If used will create disabled Actions. Defaults to $false if omitted (create Enabled Actions).

    .Example
    New-VUEMNetDrivesXML

    Description

    -----------

    Create WEM Actions from all the Drive Mappings for the current user and export this to VUEMNetDrives.xml in the current folder.

    .Example
    New-VUEMNetDrivesXML -Prefix "ITW - "

    Description

    -----------

    Create WEM Actions from all the Drive Mappings for the current user and export this to VUEMNetDrives.xml in the current folder.
    All Action names in the WEM Console are prefixed with "ITW - ", like for example "ITW - P: Powershell Scripts".

    .Example
    New-VUEMNetDrivesXML -DriveLetter "P"

    Description

    -----------

    Create a WEM Action for only the P: Drive Mapping for the current user and export this to VUEMNetDrives.xml in the current folder.

    .Example
    New-VUEMNetDrivesXML -OutputPath "C:\Temp" -OutputFileName "drives.xml"

    Description

    -----------

    Create drives.xml in c:\temp containing WEM Actions for all the Drive Mappings for the current user.

    .Example
    New-VUEMNetDrivesXML -SelfHealing Enabled -Disable

    Description

    -----------

    Create WEM Actions from all the Drive Mappings for the current user and export this to VUEMNetDrives.xml in the current folder.
    Actions are created using the Enable SelfHealing switch, and are disabled.

    .Notes
    Credentials are skipped.
#>
function New-VUEMNetDrivesXML {
    param(
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$DriveLetter,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$OutputFileName = "VUEMNetDrives.xml",
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$Prefix,
        [Parameter(Mandatory=$False,        
        ValueFromPipeline=$False)][switch]$SelfHealingEnabled,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][switch]$Disable
    )

    # check if $DriveLetter is valid
    If ($DriveLetter -and !(Test-Path "HKCU:Network\$DriveLetter")) {
        Write-Host "Cannot find '$DriveLetter' because it does not exist or is not a valid driveletter." -ForegroundColor Red
        Break
   }

   # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path $OutputPath) -or ((Get-Item $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    }

    # grab mapped drives
    $MappedDrives = @()
    
    # tell user we're processing Network Drives
    Write-Host "`nProcessing Mapped Drives" -ForegroundColor Yellow
    If ($DriveLetter) {
        $MappedDrivesRegistry = Get-ChildItem "HKCU:Network" | Where-Object {$_.PSChildName -like $DriveLetter}
    } Else {
        $MappedDrivesRegistry = Get-ChildItem "HKCU:Network"
    }
    ForEach ($MappedDrive in $MappedDrivesRegistry) {
        $MappedDrives += Get-ItemProperty "HKCU:Network\$($MappedDrive.PSChildName)"
    }

    # init VUEM action arrays
    $VUEMNetDrives = @()

    # define state
    $State = "1"
    If ($Disable) { $State = "0" }

    # define selfhealing
    $SelfHeal = "0"
    If ($SelfHealingEnabled) { $SelfHeal = "1" }

    # process inputs
    ForEach ($MappedDrive in $MappedDrives) {
        $DriveName = "$Prefix$($MappedDrive.PSChildName.ToUpper()):"
        $MappedDriveLabel = (Get-ItemProperty "HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($MappedDrive.RemotePath.Replace('\','#'))")._LabelFromReg
        If ($MappedDriveLabel) { $DriveName += " $MappedDriveLabel" }
                
        $MappedDriveName = Get-UniqueActionName -ObjectList $VUEMNetDrives -ActionName $DriveName

        $VUEMNetDrives += New-VUEMNetDriveObject -Name "$MappedDriveName" `
                                                -Description $null `
                                                -DisplayName "$MappedDriveLabel" `
                                                -TargetPath "$($MappedDrive.RemotePath)" `
                                                -SelfHealingEnabled "$SelfHealingEnabled" `
                                                -State "$State"
    }

    # output xml file
    If ($VUEMNetDrives) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMNetDrive" -ObjectList $VUEMNetDrives | Out-File $OutputPath\$OutputFileName
        Write-Host "$OutputFileName written to '$OutputPath\$OutputFileName'" -ForegroundColor Green
    }
}

<# 
    .Synopsis
    Builds an .xml file containing WEM Action definitions.

    .Description
    Builds an .xml file containing WEM Action definitions for UserDSN entries.
    This function supports multiple types of input and creates the file containing the Actions
    ready for import into WEM.

    .Link
    https://msfreaks.wordpress.com

    .Parameter Name
    Will only process the DSN defined by "Name".

    .Parameter OutputPath
    Location where the output xml file will be written to. Defaults to current folder if
    omitted.

    .Parameter OutputFileName
    The default filename is VUEMUserDSNs.xml. Use this parameter to override this if needed. 

    .Parameter SystemDSN
    If this parameter is used, the script will process SystemDSN into UserDSN Actions.

    .Parameter Prefix
    Provide a prefix string used to generate Action names (as displayed in the WEM console).

    .Parameter RunOnce
    If used will create Actions using the RunOnce parameter. Defaults to $false if omitted.

    .Parameter Disable
    If used will create disabled Actions. Defaults to $false if omitted (create Enabled Actions).

    .Example
    New-VUEMUserDSNsXML

    Description

    -----------

    Create WEM Actions from all the User DSNs for the current user and export this to VUEMUserDSNs.xml in the current folder.

    .Example
    New-VUEMUserDSNsXML -Name "WEM Database Connection" -System -RunOnce

    Description

    -----------

    Create WEM UserDSN Action from the "WEM Database Connection" for either the current user, or for the local system.
    Eventhough this might be a System DSN, the script will create a UserDSN action.
    The RunOnce parameter will be enabled for these actions.

    .Example
    New-VUEMUserDSNsXML -Prefix "ITW - "

    Description

    -----------

    Create WEM Actions from all the User DSNs for the current user and export this to VUEMUserDSNs.xml in the current folder.
    All Action names in the WEM Console are prefixed with "ITW - ", like for example "ITW - Notepad".

    .Example
    New-VUEMUserDSNsXML -OutputPath "C:\Temp" -OutputFileName "dsns.xml" -Disable

    Description

    -----------

    Create dsns.xml in c:\temp for all the User DSNs for the current user. The Actions will be disabled once imported into WEM.

    .Notes
    Seems WEM only supports UserDSN based on the "SQL Server" driver, so all DataSources based on other drivers are skipped.
    Credentials are skipped.
#>
function New-VUEMUserDSNsXML {
    param(
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$Name,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$OutputFileName = "VUEMUserDSNs.xml",
        [Parameter(Mandatory=$False,        
        ValueFromPipeline=$False)][switch]$SystemDSN,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$Prefix,
        [Parameter(Mandatory=$False,        
        ValueFromPipeline=$False)][switch]$RunOnce,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][switch]$Disable
    )

    # check if $Name is valid
    If (($Name -and $SystemDSN -and (!(Get-OdbcDsn | Where-Object {$_.Name -like "$Name"}))) `
        -or ($Name -and !($SystemDSN) -and (!(Get-OdbcDsn | Where-Object {$_.Name -like "$Name" -and $_.DsnType -eq "User"})))) {
        Write-Host "Cannot find User DSN or System DSN '$Name' because it does not exist." -ForegroundColor Red
        Break
    }

    # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path $OutputPath) -or ((Get-Item $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    }

    # grab dsns
    $dsns = @()
    
    # tell user we're processing DSNs
    Write-Host "`nProcessing DSNs" -ForegroundColor Yellow
    If ($Name -and $System) {
        $dsns = Get-OdbcDsn | Where-Object { $_.DriverName -eq "SQL Server" -and $_.Name -like "$Name" }
    } Elseif ($Name -and !$System) {
        $dsns = Get-OdbcDsn | Where-Object { $_.DriverName -eq "SQL Server" -and  $_.Name -like "$Name" -and $_.DsnType -eq "User" }
    } Elseif (!$Name -and $System) {
        $dsns = Get-OdbcDsn | Where-Object { $_.DriverName -eq "SQL Server" }
    } ElseIf (!$Name -and !$System) {
        $dsns = Get-OdbcDsn | Where-Object { $_.DriverName -eq "SQL Server" -and  $_.DsnType -eq "User" }
    }

    # init VUEM action arrays
    $VUEMUserDSNs = @()

    # define state
    $State = "1"
    If ($Disable) { $State = "0" }

    # define RunOnce
    $VUEMRunOnce = "0"
    If ($RunOnce) { $VUEMRunOnce = "1" }

    # process inputs
    ForEach ($dsn in $dsns) {

        $VUEMDSNName = Get-UniqueActionName -ObjectList $VUEMUserDSNs -ActionName "$Prefix$($dsn.Name)"

        $VUEMUserDSNs += New-VUEMUserDSNObject -Name "$VUEMDSNName" `
                                               -Description "$($dsn.Attribute.Description)" `
                                               -TargetName "$($dsn.Name)" `
                                               -TargetDriverName "$($dsn.DriverName)" `
                                               -TargetServerName "$($dsn.Attribute.Server)" `
                                               -TargetDatabaseName "$($dsn.Attribute.Database)" `
                                               -RunOnce "$VUEMRunOnce" `
                                               -State "$State"
    }

    # output xml file
    If ($VUEMUserDSNs) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMUserDSN" -ObjectList $VUEMUserDSNs | Out-File $OutputPath\$OutputFileName
        Write-Host "$OutputFileName written to '$OutputPath\$OutputFileName'" -ForegroundColor Green
    }
}

<# 
    .Synopsis
    Imports User Preference settings from GPOs and converts them to WEM Action files.

    .Description
    Imports User Preference settings from GPOs and converts them to WEM Action files.

    .Link
    https://msfreaks.wordpress.com

    .Parameter GPOBackupPath
    This is the path where the GPO Backup files are stored.
    GPO Backups are each stored in its own folder like {<GPO Backup GUID>}.
    All GPO Backups in the GPOBackupPath are processed. 

    .Parameter OutputPath
    Location where the output xml files will be written to. Defaults to current folder if
    omitted.

    .Parameter Prefix
    Provide a prefix string used to generate Action names (as displayed in the WEM console).

    .Parameter Disable
    If used will create disabled Actions. Defaults to $false if omitted (create Enabled Actions).

    .Example
    Import-VUEMActionsFromGPO -GPOBackupPath C:\GPOBackups

    Description

    -----------

    Create WEM Actions from all the user preferences as defined in all the GPOBackups found in C:\GPOBackups.
    Output is created in the current folder.

    .Example
    Import-VUEMActionsFromGPO -GPOBackupPath C:\GPOBackups -OutputPath "C:\Temp"

    Description

    -----------

    Create WEM Actions from all the user preferences as defined in all the GPOBackups found in C:\GPOBackups.
    Output is created in c:\temp folder.

    .Example
    Import-VUEMActionsFromGPO -GPOBackupPath C:\GPOBackups -Prefix "ITW - " -Disable

    Description

    -----------

    Create WEM Actions from all the user preferences as defined in all the GPOBackups found in C:\GPOBackups.
    Output is created in the current folder.
    All Action names in the WEM Console are prefixed with "ITW - ", like for example "ITW - UserDSN1", and all actions are disabled.

    .Notes
    Drive Mappings
    -----------
    If the GPO for a drive mapping has the action set to D (Delete), the drive mapping preference will be skipped.
    The Action Name will be based on the preference name and its label, if it exists.
    If the preference has the RunOne switch enabled, the SelfHealing switch will be disabled.
    If the preference has the RunOne switch disabled, the SelfHealing switch will be enabled.
    Credentials are skipped.

    Environment Variables
    -----------
    Both System and User variables from the User GPO will be processed.

    Files
    -----------
    Since WEM doesn't differentiate between files and folders where Actions are concerned, both preferences are processed into the same array of actions.
    If the GPO for a file or folder action has the action set to D (Delete), the Action Name is suffixed with " (Delete)".
    For file creation Actions the ActionType is set to 0.
    For folder creation Actions the ActionType is set to 5.
    For file or folder deletion Actions the ActionType is set to 1.

    IniFiles
    -----------
    IniFile User GPO Preferences are processed as is.

    Printer Mappings
    -----------
    If the GPO for a printer mapping has the action set to D (Delete), the printer mapping preference will be skipped.
    The Action Name will be based on the preference name and its label, if it exists.
    If the preference has the RunOne switch enabled, the SelfHealing switch will be disabled.
    If the preference has the RunOne switch disabled, the SelfHealing switch will be enabled.
    Printer Mappings are created as Map Network Printer Actions.
    Only Printer Mappings from the GPO User Preferences are processed. If you published printers to the GPO, these will be skipped.
    Credentials are skipped.

    Registry Settings
    -----------
    If the GPO for a Registry action has the action set to D (Delete), the Action Name is suffixed with " (Delete)".
    Collections are processed as individual actions, Collection names are omitted.

    DataSources
    -----------
    DataSources are processed into UserDSN actions.
    System DSN preferences will be skipped.
    Seems WEM only supports UserDSN based on the "SQL Server" driver, so all DataSources based on other drivers are skipped.
    Credentials are skipped.
#>
function Import-VUEMActionsFromGPO {
    param(
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True)][string]$GPOBackupPath,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string]$Prefix,
        [Parameter(Mandatory=$False,        
        ValueFromPipeline=$False)][switch]$Disable
    )
    
    # check if $GPOBackupPath is valid
    If ($GPOBackupPath -and !(Test-Path $GPOBackupPath)) {
        Write-Host "Cannot find path '$GPOPath' because it does not exist." -ForegroundColor Red
        Break
    }
    if ($GPOBackupPath.EndsWith("\")) { $GPOBackupPath = $GPOBackupPath.Substring(0,$GPOBackupPath.Length-1) }

    # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path $OutputPath) -or ((Get-Item $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    } Elseif ($OutputPath.EndsWith("\")) { $OutputPath = $OutputPath.Substring(0,$OutputPath.Length-1) }

    # grab GPO backups
    $GPOBackups = @()
    $GPOBackups = Get-ChildItem -Path $GPOBackupPath -Directory | Where-Object {$_.Name -like "{*}" } | Select-Object FullName

    If (!$GPOBackups) {
        Write-Host "Connot locate GPO Backups in '$GPOPath'" -ForegroundColor Red
        Break
    }

    # init VUEM action arrays
    $VUEMNetDrives = @()
    $VUEMEnvVariables = @()
    $VUEMFileSystemOps = @()
    $VUEMIniFileOps = @()
    $VUEMPrinters = @()
    $VUEMRegValues = @()
    $VUEMUserDSNs = @()

    # define state
    $State = "1"
    If ($Disabled) { $State = "0" }

    # process all GPO backups in the folder
    ForEach ($GPOBackup in $GPOBackups) {
        # set GPO Preference path
        $GPOPreferenceLocation = $GPOBackup.FullName + "\DomainSysvol\GPO\User\Preferences\"

        #region GPO Preferences - Drives
        If (Test-Path ("{0}Drives\Drives.xml" -f $GPOPreferenceLocation)) {
            write-host "Found Drives User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab Drives from Drives.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}Drives\Drives.xml" -f $GPOPreferenceLocation)
            # grab Drives where action is not set to D (Delete)
            $GPODrives = $GPOPreference.Drives.Drive | Where-Object { $_.Properties.action -notlike "D" }
            
            # convert Drives to VUEMNetDrives
            ForEach ($GPODrive in $GPODrives) {
                $GPOName = "$Prefix$($GPODrive.Name)"
                If ($GPODrive.Properties.label) { $GPOName += " $($GPODrive.Properties.label)" }
                
                $GPOSelfHealingEnabled = "1"
                If ($GPODrive.Filters.FilterRunOnce) { $GPOSelfHealingEnabled = "0" }
                
                $GPODriveName = Get-UniqueActionName -ObjectList $VUEMNetDrives -ActionName $GPOName

                $VUEMNetDrives += New-VUEMNetDriveObject -Name "$GPODriveName" `
                                                        -Description "$($GPODrive.desc)" `
                                                        -DisplayName "$($GPODrive.Properties.label)" `
                                                        -TargetPath "$($GPODrive.Properties.path)" `
                                                        -SelfHealingEnabled "$GPOSelfHealingEnabled" `
                                                        -State "$State"
            }
        }
        #endregion

        #region GPO Preferences - Environment Variables
        If (Test-Path ("{0}EnvironmentVariables\EnvironmentVariables.xml" -f $GPOPreferenceLocation)) {
            write-host "Found EnvironmentVariables User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab EnvironmentVariables from EnvironmentVariables.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}EnvironmentVariables\EnvironmentVariables.xml" -f $GPOPreferenceLocation)
            # grab EnvironmentVariables
            $GPOEnvironmentVariables = $GPOPreference.EnvironmentVariables.EnvironmentVariable
            
            # convert EnvironmentVariables  to VUEMEnvVariables
            ForEach ($GPOEnvironmentVariable in $GPOEnvironmentVariables) {
                $GPOEnvironmentVariableName = Get-UniqueActionName -ObjectList $VUEMEnvVariables -ActionName "$Prefix$($GPOEnvironmentVariable.name)"

                $GPOEnvironmentVariableType = "User"
                If ($GPOEnvironmentVariable.Properties.user -ne "1") { $GPOEnvironmentVariableType = "System" }

                $VUEMEnvVariables += New-VUEMEnvVariableObject -Name "$GPOEnvironmentVariableName" `
                                                            -Description "$($GPOEnvironmentVariable.desc)" `
                                                            -VariableName "$($GPOEnvironmentVariable.Properties.name)" `
                                                            -VariableValue "$($GPOEnvironmentVariable.Properties.value)" `
                                                            -VariableType $GPOEnvironmentVariableType `
                                                            -State "$State"
            }
        }
        #endregion

        #region GPO Preferences - Files
        If (Test-Path ("{0}Files\Files.xml" -f $GPOPreferenceLocation)) {
            write-host "Found Files User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab Files from Files.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}Files\Files.xml" -f $GPOPreferenceLocation)
            # grab Files
            $GPOFiles = $GPOPreference.Files.File
            
            # convert Files to VUEMNetFileSystemOps
            ForEach ($GPOFile in $GPOFiles) {
                $GPOName = "$Prefix$($GPOFile.Name)"

                $GPOFileRunOnce = "0"
                If ($GPOFile.Filters.FilterRunOnce) { $GPOFileRunOnce = "1" }
                
                $GPOFileAction = "0"
                $GPOFileSourcePath = $GPOFile.Properties.fromPath
                $GPOFileTargetPath = $GPOFile.Properties.targetPath
                If ($GPOFile.Properties.action -eq "D") { 
                    $GPOFileAction = "1"
                    $GPOName += " (Delete)"
                    $GPOFileSourcePath = $GPOFile.Properties.targetPath
                    $GPOFileTargetPath = $null
                }

                $GPOFileName = Get-UniqueActionName -ObjectList $VUEMFileSystemOps -ActionName $GPOName

                $VUEMFileSystemOps += New-VUEMFileSystemOpObject -Name "$GPOFileName" `
                                                                -Description "$($GPOFile.desc)" `
                                                                -SourcePath "$GPOFileSourcePath" `
                                                                -TargetPath "$GPOFileTargetPath" `
                                                                -RunOnce "$GPOFileRunOnce" `
                                                                -ActionType "$GPOFileAction" `
                                                                -State "$State"
            }
        }
        #endregion

        #region GPO Preferences - Folders
        If (Test-Path ("{0}Folders\Folders.xml" -f $GPOPreferenceLocation)) {
            write-host "Found Folders User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab Folders from Files.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}Folders\Folders.xml" -f $GPOPreferenceLocation)
            # grab Folders
            $GPOFolders = $GPOPreference.Folders.Folder
            
            # convert Folders to VUEMNetFileSystemOps
            ForEach ($GPOFolder in $GPOFolders) {
                $GPOName = "$Prefix$($GPOFolder.Name)"

                $GPOFolderRunOnce = "0"
                If ($GPOFolder.Filters.FilterRunOnce) { $GPOFolderRunOnce = "1" }

                $GPOFolderAction = "5"
                If ($GPOFolder.Properties.action -eq "D") { 
                    $GPOFolderAction = "1"
                    $GPOName += " (Delete)"
                }

                $GPOFolderName = Get-UniqueActionName -ObjectList $VUEMFileSystemOps -ActionName $GPOName

                $VUEMFileSystemOps += New-VUEMFileSystemOpObject -Name "$GPOFolderName" `
                                                                -Description "$($GPOFolder.desc)" `
                                                                -SourcePath "$($GPOFolder.Properties.path)" `
                                                                -TargetPath $null `
                                                                -RunOnce "$GPOFolderRunOnce" `
                                                                -ActionType "$GPOFolderAction" `
                                                                -State "$State"
            }
        }
        #endregion

        #region GPO Preferences - IniFiles
        If (Test-Path ("{0}IniFiles\IniFiles.xml" -f $GPOPreferenceLocation)) {
            write-host "Found IniFiles User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab IniFiles from IniFiles.xml 
            [xml]$GPOPreference = Get-Content -Path ("{0}IniFiles\IniFiles.xml" -f $GPOPreferenceLocation)
            # grab IniFiles  where action is not set to D (Delete)
            $GPOIniFiles = $GPOPreference.IniFiles.Ini | Where-Object { $_.Properties.action -notlike "D" }
            
            # convert IniFiles to VUEMNetFileSystemOps
            ForEach ($GPOIniFile in $GPOIniFiles) {
                $GPOIniFileName = Get-UniqueActionName -ObjectList $VUEMIniFileOps -ActionName "$Prefix$($GPOIniFile.Name)"

                $GPOIniFileRunOnce = "0"
                If ($GPOIniFile.Filters.FilterRunOnce) { $GPOIniFileRunOnce = "1" }

                $VUEMIniFileOps += New-VUEMIniFileOpObject -Name "$GPOIniFileName" `
                                                        -Description "$($GPOIniFile.desc)" `
                                                        -TargetPath "$($GPOIniFile.Properties.path)" `
                                                        -TargetSectionName "$($GPOIniFile.Properties.section)" `
                                                        -TargetValueName "$($GPOIniFile.Properties.property)" `
                                                        -TargetValue "$($GPOIniFile.Properties.value)" `
                                                        -RunOnce "$GPOIniFileRunOnce" `
                                                        -State "$State"
            }
        }
        #endregion
        
        #region GPO Preferences - Printers
        If (Test-Path ("{0}Printers\Printers.xml" -f $GPOPreferenceLocation)) {
            write-host "Found Printers User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab Printers from Printers.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}Printers\Printers.xml" -f $GPOPreferenceLocation)
            # grab Printers where action is not set to D (Delete)
            $GPOPrinters = $GPOPreference.Printers.SharedPrinter | Where-Object { $_.Properties.action -notlike "D" }
            
            # convert Printers to VUEMPrinters
            ForEach ($GPOPrinter in $GPOPrinters) {
                $GPOName = "$Prefix$($GPOPrinter.Name)"
                If ($GPOPrinter.Properties.label) { $GPOName += " $($GPOPrinter.Properties.label)" }

                $GPOSelfHealingEnabled = "1"
                If ($GPOPrinter.Filters.FilterRunOnce) { $GPOSelfHealingEnabled = "0" }

                $GPOPrinterName = Get-UniqueActionName -ObjectList $VUEMPrinters -ActionName $GPOName

                $VUEMPrinters += New-VUEMPrinterObject -Name "$GPOPrinterName" `
                                                    -Description "$($GPOPrinter.desc)" `
                                                    -TargetPath "$($GPOPrinter.Properties.path)" `
                                                    -SelfHealingEnabled "$GPOSelfHealingEnabled" `
                                                    -State "$State"
            }
        }
        #endregion

        #region GPO Preferences - Registry
        If (Test-Path ("{0}Registry\Registry.xml" -f $GPOPreferenceLocation)) {
            write-host "Found Registry User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab Registry from Registry.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}Registry\Registry.xml" -f $GPOPreferenceLocation)
            # grab Registry
            $GPORegistrySettings = $GPOPreference.RegistrySettings.Registry
            $GPORegistrySettings += $GPOPreference.RegistrySettings.Collection.Registry
            
            # convert RegistrySettings to VUEMRegValues
            ForEach ($GPORegistrySetting in $GPORegistrySettings) {
                $GPOName = "$Prefix$($GPORegistrySetting.Name)"

                $GPORegistrySettingRunOnce = "0"
                If ($GPORegistrySetting.Filters.FilterRunOnce) { $GPORegistrySettingRunOnce = "1" }

                $GPORegistrySettingAction = "0"
                If ($GPORegistrySetting.Properties.action -eq "D") { 
                    $GPORegistrySettingAction = "1"
                    $GPOName += " (Delete)"
                }

                $GPORegistrySettingName = Get-UniqueActionName -ObjectList $VUEMRegValues -ActionName $GPOName

                $VUEMRegValues += New-VUEMRegValueObject -Name "$GPORegistrySettingName" `
                                                        -Description "$($GPORegistrySetting.desc)" `
                                                        -TargetName "$($GPORegistrySetting.Properties.name)" `
                                                        -TargetPath "$($GPORegistrySetting.Properties.key)" `
                                                        -TargetType "$($GPORegistrySetting.Properties.type)" `
                                                        -TargetValue "$($GPORegistrySetting.Properties.value)" `
                                                        -RunOnce "$GPORegistrySettingRunOnce" `
                                                        -ActionType "$GPORegistrySettingAction" `
                                                        -State "$State"
            }
        }
        #endregion

        #region GPO Preferences - DataSources
        If (Test-Path ("{0}DataSources\DataSources.xml" -f $GPOPreferenceLocation)) {
            write-host "Found DataSources User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab DataSources from DataSources.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}DataSources\DataSources.xml" -f $GPOPreferenceLocation)
            # grab DataSources where action is not set to D (Delete), driver equals "SQL Server" and User DSN only
            $GPODataSources = $GPOPreference.DataSources.DataSource | Where-Object { $_.Properties.action -notlike "D" -and $_.Properties.driver -eq "SQL Server" -and $_.Properties.userDSN -eq "1" }
            
            # convert RegistrySettings to VUEMRegValues
            ForEach ($GPODataSource in $GPODataSources) {
                $GPODataSourceName = Get-UniqueActionName -ObjectList $VUEMNetDrives -ActionName "$Prefix$($GPODataSource.Name)"

                $GPODataSourceRunOnce = "0"
                If ($GPODataSource.Filters.FilterRunOnce) { $GPODataSourceRunOnce = "1" }

                $GPODataSourceServerName = ($GPODataSource.Properties.Attributes.Attribute | Where-Object {$_.name -eq "SERVER"}).value
                $GPODataSourceDatabaseName = ($GPODataSource.Properties.Attributes.Attribute | Where-Object {$_.name -eq "DATABASE"}).value

                $VUEMUserDSNs += New-VUEMUserDSNObject -Name "$GPODataSourceName" `
                                                    -Description "$($GPODataSource.Properties.description)" `
                                                    -TargetName "$($GPODataSource.Properties.dsn)" `
                                                    -TargetDriverName "$($GPODataSource.Properties.driver)" `
                                                    -TargetServerName "$GPODataSourceServerName" `
                                                    -TargetDatabaseName "$GPODataSourceDatabaseName" `
                                                    -RunOnce "$GPODataSourceRunOnce" `
                                                    -State "$State"
            }
        }
        #endregion

    }

    # output xml files
    If ($VUEMNetDrives) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMNetDrive" -ObjectList $VUEMNetDrives | Out-File $OutputPath\VUEMNetDrives.xml
        Write-Host "VUEMNetDrives.xml written to '$OutputPath\VUEMNetDrives.xml'" -ForegroundColor Green
    }
    If ($VUEMEnvVariables) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMEnvVariable" -ObjectList $VUEMEnvVariables | Out-File $OutputPath\VUEMEnvVariables.xml
        Write-Host "VUEMEnvVariables.xml written to '$OutputPath\VUEMEnvVariables.xml'" -ForegroundColor Green
    }
    If ($VUEMFileSystemOps) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMFileSystemOp" -ObjectList $VUEMFileSystemOps | Out-File $OutputPath\VUEMFileSystemOps.xml
        Write-Host "VUEMFileSystemOps.xml written to '$OutputPath\VUEMFileSystemOps.xml'" -ForegroundColor Green
    }
    If ($VUEMIniFileOps) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMIniFileOp" -ObjectList $VUEMIniFileOps | Out-File $OutputPath\VUEMIniFileOps.xml
        Write-Host "VUEMIniFileOps.xml written to '$OutputPath\VUEMIniFileOps.xml'" -ForegroundColor Green
    }
    If ($VUEMPrinters) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMPrinter" -ObjectList $VUEMPrinters | Out-File $OutputPath\VUEMPrinters.xml
        Write-Host "VUEMPrinters.xml written to '$OutputPath\VUEMPrinters.xml'" -ForegroundColor Green
    }
    If ($VUEMRegValues) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMRegValue" -ObjectList $VUEMRegValues | Out-File $OutputPath\VUEMRegValues.xml
        Write-Host "VUEMRegValues.xml written to '$OutputPath\VUEMRegValues.xml'" -ForegroundColor Green
    }
    If ($VUEMUserDSNs) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMUserDSN" -ObjectList $VUEMUserDSNs | Out-File $OutputPath\VUEMUserDSNs.xml
        Write-Host "VUEMUserDSNs.xml written to '$OutputPath\VUEMUserDSNs.xml'" -ForegroundColor Green
    }
}

#region Helper Functions (will not be exposed when module is loaded)
<#
 .SYNOPSIS
  Helper function to grab an array of files to process.
#>
function Get-FilesToProcess{
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )

    If ((Get-Item $Path) -is [System.IO.DirectoryInfo]) {
        # what to do if $Path is a folder
        $relativePath = $Path
        If (!$Path.EndsWith("\*")) { 
            $Path += "\*"
        } Else {
            $relativePath = $Path.Replace("\*","")
        }

        # build filter
        $Filter = "*.lnk"
        If ($FileTypes) { 
            $Filter = @()
            ForEach ($FileType in ($FileTypes.Split(","))) {
                $Filter += "*." + $($FileType.Replace("*.","").Trim())
            }
        }

        # to recurse or not to recurse
        If ($Recurse){
            $files = Get-ChildItem -Path $Path -Include $Filter -Recurse | Select-Object Name, @{ n = 'BaseName'; e = { $_.BaseName } }, @{ n = 'FolderName'; e = { Convert-Path $_.PSParentPath } }, @{ n = 'RelativePath'; e = { (Convert-Path $_.PSParentPath).Replace($relativePath,"") } }, FullName, @{ n = 'Extension'; e = { $_.Extension } } | Where-Object { !$_.PSIsContainer }
        } Else {
            $files = Get-ChildItem -Path $Path -Include $Filter | Select-Object Name, @{ n = 'BaseName'; e = { $_.BaseName } }, @{ n = 'FolderName'; e = { Convert-Path $_.PSParentPath } }, @{ n = 'RelativePath'; e = { (Convert-Path $_.PSParentPath).Replace($relativePath,"") } }, FullName, @{ n = 'Extension'; e = { $_.Extension } } | Where-Object { !$_.PSIsContainer }
        }

    } Else {
        # what to do if $Path is a single file
        $files = Get-ChildItem -Path $Path -File
    }

    Return $files
}

<#
 .SYNOPSIS
  Helper function to grab an icon into Base64 encoded string.
#>
function Get-IconToBase64{
    Param(
        [Parameter(Mandatory=$true)][object]$Icon
    )

	$stream = New-Object System.IO.MemoryStream
	$bmp = $Icon.ToBitmap()
	$bmp.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)

    Return ([System.Convert]::ToBase64String($stream.ToArray()))
}

<#
 .SYNOPSIS
  Helper function to create VUEMApplication object
#>
function New-VUEMApplicationObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$DisplayName,
        [string]$StartMenuTarget,
        [string]$TargetPath,
        [string]$Parameters,
        [string]$WorkingDirectory,
        [string]$Hotkey,
        [string]$IconLocation,
        [string]$IconStream,
        [string]$SelfHealingEnabled,
        [string]$State
    )
    
    return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'Displayname' = $DisplayName
        'StartMenuTarget' = $StartMenuTarget
        'TargetPath' = $TargetPath
        'Parameters' = $Parameters
        'WorkingDirectory' = $WorkingDirectory
        'WindowStyle' = "Normal"
        'Hotkey' = $Hotkey
        'IconLocation' = $IconLocation
        'IconIndex' = "0"
        'IconStream' = $IconStream
        'Reserved01' = '<?xml version="1.0" encoding="utf-8"?><ArrayOfVUEMActionAdvancedOption xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><VUEMActionAdvancedOption><Name>SelfHealingEnabled</Name><Value>'+$SelfHealingEnabled+'</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>EnforceIconLocation</Name><Value>0</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>EnforcedIconXValue</Name><Value>0</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>EnforcedIconYValue</Name><Value>0</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>DoNotShowInSelfService</Name><Value>0</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>CreateShortcutInUserFavoritesFolder</Name><Value>0</Value></VUEMActionAdvancedOption></ArrayOfVUEMActionAdvancedOption>'
        'ActionType' = "0"
        'State' = $State
    }
    # Action type 0 = create application shortcut
}

<#
 .SYNOPSIS
  Helper function to create VUEMNetDrive object
#>
function New-VUEMNetDriveObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$DisplayName,
        [string]$TargetPath,
        [string]$SelfHealingEnabled,
        [string]$State
    )
    
    return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'Displayname' = $DisplayName
        'TargetPath' = $TargetPath
        'Reserved01' = '<?xml version="1.0" encoding="utf-8"?><ArrayOfVUEMActionAdvancedOption xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><VUEMActionAdvancedOption><Name>SelfHealingEnabled</Name><Value>'+$SelfHealingEnabled+'</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>SetAsHomeDriveEnabled</Name><Value>0</Value></VUEMActionAdvancedOption></ArrayOfVUEMActionAdvancedOption>'
        'ActionType' = "0"
        'UseExtCredentials' = "0"
        'Extlogin' = $null
        'ExtPassword' = "Uah1meBsLIw="
        'State' = $State
    }
    # Action type 0 = map drive
}

<#
 .SYNOPSIS
  Helper function to create VUEMEnvVariable object
#>
function New-VUEMEnvVariableObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$VariableName,
        [string]$VariableValue,
        [string]$VariableType,
        [string]$State
    )

    return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'VariableName' = $VariableName
        'VariableValue' = $VariableValue
        'VariableType' = $VariableType
        'Reserved01' = '<?xml version="1.0" encoding="utf-8"?><ArrayOfVUEMActionAdvancedOption xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><VUEMActionAdvancedOption><Name>ExecOrder</Name><Value>0</Value></VUEMActionAdvancedOption></ArrayOfVUEMActionAdvancedOption>'
        'ActionType' = "0"
        'State' = $State
    }
    # Action type 0 = create environment variable 
}

<#
 .SYNOPSIS
  Helper function to create VUEMFileSystemOp object
#>
function New-VUEMFileSystemOpObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$SourcePath,
        [string]$TargetPath,
        [string]$RunOnce,
        [string]$ActionType,
        [string]$State
    )

    return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'SourcePath' = $SourcePath
        'TargetPath' = $TargetPath
        'RunOnce' = $RunOnce
        'TargetOverwrite' = "1"
        'Reserved01' = '<?xml version="1.0" encoding="utf-8"?><ArrayOfVUEMActionAdvancedOption xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><VUEMActionAdvancedOption><Name>ExecOrder</Name><Value>0</Value></VUEMActionAdvancedOption></ArrayOfVUEMActionAdvancedOption>'
        'ActionType' = $ActionType
        'State' = $State
    }
    # Action type 0 = copy file / folder
    # Action type 1 = remove file / folder
    # Action type 5 = create folder
    # Action type 7 = remove folder content
}

<#
 .SYNOPSIS
  Helper function to create VUEMIniFileOp object
#>
function New-VUEMIniFileOpObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$TargetPath,
        [string]$TargetSectionName,
        [string]$TargetValueName,
        [string]$TargetValue,
        [string]$RunOnce,
        [string]$State        
    )

    return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'TargetPath' = $TargetPath
        'TargetSectionName' = $TargetSectionName
        'TargetValueName' = $TargetValueName
        'TargetValue' = $TargetValue
        'RunOnce' = $RunOnce
        'Reserved01' = $null
        'ActionType' = "0"
        'State' = $State
    }
    # Action type 0 = write ini value 
}

<#
 .SYNOPSIS
  Helper function to create VUEMPrinter object
#>
function New-VUEMPrinterObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$TargetPath,
        [string]$SelfHealingEnabled,
        [string]$State
    )
    
    return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'DisplayName' = $null
        'TargetPath' = $TargetPath
        'Reserved01' = '<?xml version="1.0" encoding="utf-8"?><ArrayOfVUEMActionAdvancedOption xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><VUEMActionAdvancedOption><Name>SelfHealingEnabled</Name><Value>'+$SelfHealingEnabled+'</Value></VUEMActionAdvancedOption></ArrayOfVUEMActionAdvancedOption>'
        'ActionType' = "0"
        'UseExtCredentials' = "0"
        'Extlogin' = $null
        'ExtPassword' = "Uah1meBsLIw="
        'State' = $State
    }
    # Action type 0 = map drive 
}

<#
 .SYNOPSIS
  Helper function to create VUEMRegValue object
#>
function New-VUEMRegValueObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$TargetName,
        [string]$TargetPath,
        [string]$TargetType,
        [string]$TargetValue,
        [string]$RunOnce,
        [String]$ActionType,
        [string]$State
    )
    
    If(!$TargetName) { $TargetName = "(Default)" }

    return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'TargetRoot' = $null
        'TargetName' = $TargetName
        'TargetPath' = $TargetPath
        'TargetType' = $TargetType
        'TargetValue' = $TargetValue
        'RunOnce' = $RunOnce
        'Reserved01' = $null
        'ActionType' = $ActionType
        'State' = $State
    }
    # Action type 0 = write reg value 
    # Action type 1 = delete reg value 
}

<#
 .SYNOPSIS
  Helper function to create VUEMUserDSN object
#>
function New-VUEMUserDSNObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$TargetName,
        [string]$TargetDriverName,
        [string]$TargetServerName,
        [string]$TargetDatabaseName,
        [string]$RunOnce,
        [string]$State
    )

    return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'TargetName' = $TargetName
        'TargetDriverName' = $TargetDriverName
        'TargetServerName' = $TargetServerName
        'TargetDatabaseName' = $TargetDatabaseName
        'RunOnce' = $RunOnce
        'Reserved01' = $null
        'ActionType' = "0"
        'UseExtCredentials' = "0"
        'Extlogin' = $null
        'ExtPassword' = "Uah1meBsLIw="
        'State' = $State
    }
    # Action type 0 = create / edit DSN 
}

<#
 .SYNOPSIS
  Helper function for creating a VUEM xml file
#>
function New-VUEMXmlFile {
    param(
        [string]$VUEMIdentifier,
        [psobject[]]$ObjectList
    )

    # create a base XML file structure
    $StringWriter = New-Object System.IO.StringWriter
    $xmlWriter = New-Object System.Xml.XmlTextWriter $StringWriter
    $xmlWriter.Formatting = "indented"
    $xmlWriter.WriteProcessingInstruction("xml", "version='1.0' encoding='UTF-8'");
    
    # write xml root node
    $xmlWriter.WriteStartElement("ArrayOf$VUEMIdentifier")
    $xmlWriter.WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
    $xmlWriter.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");

    ForEach ($object in $ObjectList) {
        # write xml child node
        $xmlWriter.WriteStartElement("$VUEMIdentifier")

        ForEach ($p in $object.PSObject.Properties) {
            $xmlWriter.WriteElementString($p.Name,$p.Value)
        }

        $xmlWriter.WriteEndElement()
    }

    # finish the XML Document
    $xmlWriter.WriteEndElement()
    $xmlWriter.Finalize
    $StringWriter.Flush()

    # return xml
    Return $StringWriter.ToString()
}

<#
 .SYNOPSIS
  Helper function to ensure a unique action name
#>
function Get-UniqueActionName {
    param(
        [psobject[]]$ObjectList,
        [string]$ActionName,
        [int]$i=1
    )

    $SearchName = $ActionName
    If ($i -gt 1) { $SearchName += " ($($i.ToString()))" }

    If (($ObjectList | Where-Object {$_.Name -eq $SearchName})) {
        $i++
        Return (Get-UniqueActionName -ObjectList $ObjectList -ActionName $ActionName -i $i)
    } Else {
        Return $SearchName
    }
}
#endregion

# expose functions
Export-ModuleMember -Function New-VUEMApplicationsXML
Export-ModuleMember -Function New-VUEMNetDrivesXML
Export-ModuleMember -Function New-VUEMUserDSNsXML
Export-ModuleMember -Function Import-VUEMActionsFromGPO
