#
# Citrix.Wem.Version = "1.1.1"
#

<# 
    .Synopsis
    Imports User Preference settings from GPOs and converts them to WEM Action files.

    .Description
    Imports User Preference settings from GPOs and converts them to WEM Action files. This function will process
    the following Action types:
    VUEMEnvVariables from GPO Preference Environment Variables
    VUEMExtTasks from GPO Policy Run these programs at user logon and GPO User Logon Scripts
    VUEMFileSystemOps from GPO Preference Files and GPO Preference Folders
    VUEMIniFileOps from GPO Preference Ini Files
    VUEMNetDrives from GPO Preference Drive Mappings
    VUEMPrinters from GPO Preference Printer Mappings and GPO Deployed Printers
    VUEMRegValues from GPO Preference Registry Settings
    VUEMUserDSNs from GPO Preference Data Sources

    .Link
    https://msfreaks.wordpress.com

    .Parameter GPOBackupPath
    This is the path where the GPO Backup files are stored.
    GPO Backups are each stored in its own folder like {<GPO Backup GUID>}.
    All GPO Backups in the GPOBackupPath are processed. 

    .Parameter OutputPath
    Location where the output xml files will be written to. Defaults to current folder if
    omitted.
    Filter export to a csv file if indicated by the parameter is also stored in this path.
    Embedded Logon Scripts and related files are exported to a subfolder in this path.
    The subfolder will be created if it does not exist.

    .Parameter Prefix
    Provide a prefix string used to generate Action names (as displayed in the WEM console).

    .Parameter SelfHealingEnabled
    If used will override any RunOnce settings found in GPO Backup files.
    WEM Actions that have the SelfHealing option will have the SelfHealing option enabled.
    WEM Actions that have the RunOnce option will have the RunOnce option disabled.

    .Parameter OverrideEmptyDescription
    If used will generate a description based on the Action name, but only if a description is not found during
    processing.

    .Parameter Disable
    If used will create disabled Actions. Defaults to $False if omitted (create Enabled Actions).

    .Parameter ExportFilters
    If used will export all Filters found in the GPO Backups while processing the User Preferences to
    'GPOFilters.csv' in the output location.

    .Example
    Import-VUEMActionsFromGpo -GPOBackupPath C:\GPOBackups

    Description

    -----------

    Create WEM Actions from all the user preferences as defined in all the GPOBackups found in C:\GPOBackups.
    Output is created in the current folder.

    .Example
    Import-VUEMActionsFromGpo -GPOBackupPath C:\GPOBackups -OutputPath "C:\Temp"

    Description

    -----------

    Create WEM Actions from all the user preferences as defined in all the GPOBackups found in C:\GPOBackups.
    Output is created in c:\temp folder. A subfolder in c:\temp is created to store Embedded Logon Scripts,
    if found during processing.

    .Example
    Import-VUEMActionsFromGpo -GPOBackupPath C:\GPOBackups -Prefix "ITW - " -Disable

    Description

    -----------

    Create WEM Actions from all the user preferences as defined in all the GPOBackups found in C:\GPOBackups.
    Output is created in the current folder.
    All Action names in the WEM Console are prefixed with "ITW - ", like for example "ITW - UserDSN1", and all
    actions are disabled.

    .Example
    Import-VUEMActionsFromGpo -GPOBackupPath C:\GPOBackups -SelfHealingEnabled -ExportFilters

    Description

    -----------

    Create WEM Actions from all the user preferences as defined in all the GPOBackups found in C:\GPOBackups.
    Output is created in the current folder.
    All Actions will have either SelfHealing set to enabled, or RunOnce set to disabled, because -SelfHealingEnabled
    was specified.
    If any GPO Filters were found during processing, they will be saved as GPOFilters.csv in the current folder.

    .Example
    Import-VUEMActionsFromGpo -OverrideEmptyDescription

    Description

    -----------

    Create WEM Actions from all the user preferences as defined in all the GPOBackups found in C:\GPOBackups.
    Output is created in the current folder.
    If no Description was found during processing, the Action name is used as description.

    .Notes
    Author:  Arjan Mensch
    Version: 1.1.1

    DataSources (VUEMUserDSNs)
    -----------
    DataSources are processed into UserDSN actions.
    System DSN preferences will be processed as User DSNs.
    WEM only supports UserDSN based on the "SQL Server" driver. All DataSources based on other drivers are skipped.
    Credentials are skipped.

    Drive Mappings (VUEMNetDrives)
    -----------
    If the GPO for a drive mapping has the action set to D (Delete), the drive mapping preference will be skipped.
    The Action Name will be based on TargetPath settings (which is a UNC path).
    If the preference has the RunOne switch enabled, the SelfHealing switch will be disabled.
    If the preference has the RunOne switch disabled, the SelfHealing switch will be enabled.
    Credentials are skipped.

    Environment Variables (VUEMEnvVariables)
    -----------
    Both System and User variables from the User GPO will be processed.

    Files / Folders (VUEMFileSystemOps)
    -----------
    Since WEM doesn't differentiate between files and folders where Actions are concerned, both preferences are
    processed into the same array of actions.
    If the GPO for a file or folder action has the action set to D (Delete), the Action Name is suffixed
    with " (Delete)".
    For file creation Actions the ActionType is set to 0.
    For folder creation Actions the ActionType is set to 5.
    For file or folder deletion Actions the ActionType is set to 1.

    IniFiles (VUEMIniFileOps)
    -----------
    IniFile User GPO Preferences are processed as is.

    Printer Mappings (VUEMPrinters)
    -----------
    If the GPO for a printer mapping has the action set to D (Delete), the printer mapping preference will
    be skipped.
    The Action Name will be based on TargetPath settings (which is a UNC path).
    If the preference has the RunOne switch enabled, the SelfHealing switch will be disabled.
    If the preference has the RunOne switch disabled, the SelfHealing switch will be enabled.
    Printer Mappings are created as Map Network Printer Actions.
    Credentials are skipped.
    Printer Mappings that were Deployed to the GPO are only processed if the GPOs are in English.

    Registry Settings (VUEMRegValues)
    -----------
    If the GPO for a Registry action has the action set to D (Delete), the Action Name is suffixed with " (Delete)".
    The Action Name will be based on Registry path, suffixed with the value name.
    Registry actions in Collections are processed as individual actions, Collection names are omitted.
    Since WEM does not support REG_BINARY settings, these will be skipped. Use -Verbose to output the names and keys
    of skipped items.

    Run these programs at user logon / User Logon Scripts (VUEMExtTasks)
    -----------
    Default values for these actions: 30s TimeOut, WaitForFinish enabled, RunHidden enabled, RunOnce disabled.
    Embedded Logon Scripts are extracted from the GPOs, along with all other files found in that location.
    External Taskes based on embedded scripts are always Disabled, and have their name prefixed
    with "[NEEDS FILE LOCATION]"
    Run these programs at user logon / User Logon Scripts are only processed if the GPOs are in English.
    This is the only action where a Computer Policy setting is also processed.
#>
Function Import-VUEMActionsFromGpo {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,Position=1)]
        [string]$GPOBackupPath,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$Prefix = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$SelfHealingEnabled = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$OverrideEmptyDescription = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$Disable = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$ExportFilters = $False
    )
    Write-Verbose "$(Get-Date -Format G)"
    Write-Verbose "Function: $($MyInvocation.MyCommand)"

    # check if $GPOBackupPath is valid
    If ($GPOBackupPath -and !(Test-Path -Path $GPOBackupPath)) {
        Write-Host "Cannot find path '$GPOBackupPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    }
    if ($GPOBackupPath.EndsWith("\")) { $GPOBackupPath = $GPOBackupPath.Substring(0,$GPOBackupPath.Length-1) }
    
    # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path -Path $OutputPath) -or ((Get-Item -Path $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    } Elseif ($OutputPath.EndsWith("\")) { $OutputPath = $OutputPath.Substring(0,$OutputPath.Length-1) }

    # grab GPO backups
    $GPOBackups = @()
    $GPOBackups = Get-ChildItem -Path $GPOBackupPath -Directory | Where-Object { $_.Name -like "{*}" } | Select-Object FullName

    If (!$GPOBackups) {
        Write-Host "Cannot locate GPO Backups in '$GPOBackupPath'" -ForegroundColor Red
        Break
    }

    # Verbose output
    Write-Verbose "Using GPOBackupPath '$GPOBackupPath'"
    Write-Verbose "Using OutputPath '$OutputPath'"
    Write-Verbose "Using Prefix '$Prefix'"
    Write-Verbose "Using SelfHealingEnabled: $SelfHealingEnabled"
    Write-Verbose "Using OverrideEmptyDescription: $OverrideEmptyDescription"
    Write-Verbose "Using Disable: $Disable"
    Write-Verbose "Using ExportFilters: $ExportFilters"
    Write-Verbose "Found $($GPOBackups.Count) GPO Backups in '$GPOBackupPath'"

    # init VUEM action arrays
    $VUEMNetDrives = @()
    $VUEMEnvVariables = @()
    $VUEMExtTasks = @()
    $VUEMFileSystemOps = @()
    $VUEMIniFileOps = @()
    $VUEMPrinters = @()
    $VUEMRegValues = @()
    $VUEMUserDSNs = @()

    # init Filter array
    $GPOFilters = @()
    $GPOScriptFiles = @()

    # define state
    $State = "1"
    If ($Disable) { $State = "0" }

    # process all GPO backups in the folder
    ForEach ($GPOBackup in $GPOBackups) {
        # set GPO Paths path
        $GPOPreferenceLocation = $GPOBackup.FullName + "\DomainSysvol\GPO\User\Preferences\"
        Write-Verbose "Processing '$($GPOBackup.FullName)'"

        # get gpreport.xml
        [xml]$GPOReport = ""
        If (Test-Path -Path ("{0}\gpreport.xml" -f $GPOBackup.FullName)) {
            [xml]$GPOReport = Get-Content -Path ("{0}\gpreport.xml" -f $GPOBackup.FullName)
        }

        #region GPO Preferences - Drives
        If (Test-Path -Path ("{0}Drives\Drives.xml" -f $GPOPreferenceLocation)) {
            Write-Host "Found Drives User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab Drives from Drives.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}Drives\Drives.xml" -f $GPOPreferenceLocation)
            # grab Drives where action is not set to D (Delete)
            $GPODrives = $GPOPreference.Drives.Drive | Where-Object { $_.Properties.action -notlike "D" }

            # convert Drives to VUEMNetDrives
            ForEach ($GPODrive in $GPODrives) {
                $GPOName = "$Prefix$($GPODrive.Properties.path)"
                If ($GPODrive.Properties.label) { $GPOName += " ($($GPODrive.Properties.label))" }
                
                $Description = "$($GPODrive.desc)".Trim()
                If (!$Description -and $OverrideEmptyDescription) { $Description = $GPOName }

                $GPOSelfHealingEnabled = "1"
                If ($GPODrive.Filters.FilterRunOnce -and (!$SelfHealingEnabled)) { $GPOSelfHealingEnabled = "0" }
                
                $GPODriveName = Get-UniqueActionName -ObjectList $VUEMNetDrives -ActionName $GPOName

                $VUEMNetDrive = New-VUEMNetDriveObject -Name "$GPODriveName" `
                                                       -Description "$Description" `
                                                       -DisplayName "$($GPODrive.Properties.label)" `
                                                       -TargetPath "$($GPODrive.Properties.path)" `
                                                       -SelfHealingEnabled "$GPOSelfHealingEnabled" `
                                                       -State "$State" `
                                                       -ObjectList $VUEMNetDrives

                If ($VUEMNetDrive) { 
                    # add new object to array
                    $VUEMNetDrives += $VUEMNetDrive
                
                    # grab GPO Filters for this new object
                    ForEach ($Filter in $GPODrive.Filters) {
                        $GPOFilter = New-GPOFilterObject -Name "$GPODriveName" -Filter $Filter -ActionType "NetDrive"
                        If ($GPOFilter) { $GPOFilters += $GPOFilter }
                    }
                }
            }
        }
        #endregion

        #region GPO Preferences - Environment Variables
        If (Test-Path -Path ("{0}EnvironmentVariables\EnvironmentVariables.xml" -f $GPOPreferenceLocation)) {
            Write-Host "Found EnvironmentVariables User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab EnvironmentVariables from EnvironmentVariables.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}EnvironmentVariables\EnvironmentVariables.xml" -f $GPOPreferenceLocation)
            # grab EnvironmentVariables
            $GPOEnvironmentVariables = $GPOPreference.EnvironmentVariables.EnvironmentVariable
           
            # convert EnvironmentVariables  to VUEMEnvVariables
            ForEach ($GPOEnvironmentVariable in $GPOEnvironmentVariables) {
                $GPOEnvironmentVariableName = Get-UniqueActionName -ObjectList $VUEMEnvVariables -ActionName "$Prefix$($GPOEnvironmentVariable.name)"

                $GPOEnvironmentVariableType = "User"
                If ($GPOEnvironmentVariable.Properties.user -ne "1") { $GPOEnvironmentVariableType = "System" }

                $Description = "$($GPOEnvironmentVariable.desc)".Trim()
                If (!$Description -and $OverrideEmptyDescription) { $Description = "$($GPOEnvironmentVariable.name)" }

                $VUEMEnvVariable = New-VUEMEnvVariableObject -Name "$GPOEnvironmentVariableName" `
                                                             -Description "$Description" `
                                                             -VariableName "$($GPOEnvironmentVariable.Properties.name)" `
                                                             -VariableValue "$($GPOEnvironmentVariable.Properties.value)" `
                                                             -VariableType $GPOEnvironmentVariableType `
                                                             -State "$State" `
                                                             -ObjectList $VUEMEnvVariables
                
                If ($VUEMEnvVariable) { 
                    # add new object to array
                    $VUEMEnvVariables += $VUEMEnvVariable
                
                    # grab GPO Filters for this new object
                    ForEach ($Filter in $GPODrive.Filters) {
                        $GPOFilter = New-GPOFilterObject -Name "$GPOEnvironmentVariableName" -Filter $Filter -ActionType "EnvVariable"
                        If ($GPOFilter) { $GPOFilters += $GPOFilter }
                    }
                }
            }
        }
        #endregion

        #region GPO Preferences - Files
        If (Test-Path -Path ("{0}Files\Files.xml" -f $GPOPreferenceLocation)) {
            Write-Host "Found Files User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab Files from Files.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}Files\Files.xml" -f $GPOPreferenceLocation)
            # grab Files
            $GPOFiles = $GPOPreference.Files.File
            
            # convert Files to VUEMNetFileSystemOps
            ForEach ($GPOFile in $GPOFiles) {
                $GPOName = "$Prefix$($GPOFile.Name)"

                $Description = "$($GPOFile.desc)".Trim()
                If (!$Description -and $OverrideEmptyDescription) { $Description = "$($GPOFile.Name)" }

                $GPOFileRunOnce = "0"
                If ($GPOFile.Filters.FilterRunOnce  -and (!$SelfHealingEnabled)) { $GPOFileRunOnce = "1" }
                
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

                $VUEMFileSystemOp = New-VUEMFileSystemOpObject -Name "$GPOFileName" `
                                                               -Description "$Description" `
                                                               -SourcePath "$GPOFileSourcePath" `
                                                               -TargetPath "$GPOFileTargetPath" `
                                                               -RunOnce "$GPOFileRunOnce" `
                                                               -ActionType "$GPOFileAction" `
                                                               -State "$State" `
                                                               -ObjectList $VUEMFileSystemOps

                If ($VUEMFileSystemOp) { 
                    # add new object to array
                    $VUEMFileSystemOps += $VUEMFileSystemOp

                    # grab GPO Filters for this new object
                    ForEach ($Filter in $GPODrive.Filters) {
                        $GPOFilter = New-GPOFilterObject -Name "$GPOFileName" -Filter $Filter -ActionType "FileSystemOp"
                        If ($GPOFilter) { $GPOFilters += $GPOFilter }
                    }
                }
            }
        }
        #endregion

        #region GPO Preferences - Folders
        If (Test-Path -Path ("{0}Folders\Folders.xml" -f $GPOPreferenceLocation)) {
            Write-Host "Found Folders User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab Folders from Files.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}Folders\Folders.xml" -f $GPOPreferenceLocation)
            # grab Folders
            $GPOFolders = $GPOPreference.Folders.Folder
            
            # convert Folders to VUEMNetFileSystemOps
            ForEach ($GPOFolder in $GPOFolders) {
                $GPOName = "$Prefix$($GPOFolder.Name)"

                $Description = "$($GPOFolder.desc)".Trim()
                If (!$Description -and $OverrideEmptyDescription) { $Description = "$($GPOFolder.Name)" }

                $GPOFolderRunOnce = "0"
                If ($GPOFolder.Filters.FilterRunOnce -and (!$SelfHealingEnabled)) { $GPOFolderRunOnce = "1" }

                $GPOFolderAction = "5"
                If ($GPOFolder.Properties.action -eq "D") { 
                    $GPOFolderAction = "1"
                    $GPOName += " (Delete)"
                }

                $GPOFolderName = Get-UniqueActionName -ObjectList $VUEMFileSystemOps -ActionName $GPOName

                $VUEMFileSystemOp = New-VUEMFileSystemOpObject -Name "$GPOFolderName" `
                                                               -Description "$Description" `
                                                               -SourcePath "$($GPOFolder.Properties.path)" `
                                                               -TargetPath $null `
                                                               -RunOnce "$GPOFolderRunOnce" `
                                                               -ActionType "$GPOFolderAction" `
                                                               -State "$State" `
                                                               -ObjectList $VUEMFileSystemOps

                If ($VUEMFileSystemOp) { 
                    # add new object to array
                    $VUEMFileSystemOps += $VUEMFileSystemOp
                
                    # grab GPO Filters for this new object
                    ForEach ($Filter in $GPODrive.Filters) {
                        $GPOFilter = New-GPOFilterObject -Name "$GPOFolderName" -Filter $Filter -ActionType "FileSystemOp"
                        If ($GPOFilter) { $GPOFilters += $GPOFilter }
                    }
                }
            }
        }
        #endregion

        #region GPO Preferences - IniFiles
        If (Test-Path -Path ("{0}IniFiles\IniFiles.xml" -f $GPOPreferenceLocation)) {
            Write-Host "Found IniFiles User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab IniFiles from IniFiles.xml 
            [xml]$GPOPreference = Get-Content -Path ("{0}IniFiles\IniFiles.xml" -f $GPOPreferenceLocation)
            # grab IniFiles  where action is not set to D (Delete)
            $GPOIniFiles = $GPOPreference.IniFiles.Ini | Where-Object { $_.Properties.action -notlike "D" }
            
            # convert IniFiles to VUEMNetFileSystemOps
            ForEach ($GPOIniFile in $GPOIniFiles) {
                $GPOIniFileName = Get-UniqueActionName -ObjectList $VUEMIniFileOps -ActionName "$Prefix$($GPOIniFile.Name)"

                $Description = "$($GPOIniFile.desc)".Trim()
                If (!$Description -and $OverrideEmptyDescription) { $Description = "$($GPOIniFile.Name)" }

                $GPOIniFileRunOnce = "0"
                If ($GPOIniFile.Filters.FilterRunOnce -and (!$SelfHealingEnabled)) { $GPOIniFileRunOnce = "1" }

                $VUEMIniFileOp = New-VUEMIniFileOpObject -Name "$GPOIniFileName" `
                                                         -Description "$Description" `
                                                         -TargetPath "$($GPOIniFile.Properties.path)" `
                                                         -TargetSectionName "$($GPOIniFile.Properties.section)" `
                                                         -TargetValueName "$($GPOIniFile.Properties.property)" `
                                                         -TargetValue "$($GPOIniFile.Properties.value)" `
                                                         -RunOnce "$GPOIniFileRunOnce" `
                                                         -State "$State" `
                                                         -ObjectList $VUEMIniFileOps

                If ($VUEMIniFileOp) { 
                    # add new object to array
                    $VUEMIniFileOps += $VUEMIniFileOp
                
                    # grab GPO Filters for this new object
                    ForEach ($Filter in $GPODrive.Filters) {
                        $GPOFilter = New-GPOFilterObject -Name "$GPOIniFileName" -Filter $Filter -ActionType "IniFileOp"
                        If ($GPOFilter) { $GPOFilters += $GPOFilter }
                    }
                }
            }
        }
        #endregion
        
        #region GPO Preferences - Printers
        If (Test-Path -Path ("{0}Printers\Printers.xml" -f $GPOPreferenceLocation)) {
            Write-Host "Found Printers User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab Printers from Printers.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}Printers\Printers.xml" -f $GPOPreferenceLocation)
            # grab Printers where action is not set to D (Delete)
            $GPOPrinters = $GPOPreference.Printers.SharedPrinter | Where-Object { $_.Properties.action -notlike "D" }
            
            # convert Printers to VUEMPrinters
            ForEach ($GPOPrinter in $GPOPrinters) {
                $GPOName = "$Prefix$($GPOPrinter.Properties.path)"
                If ($GPOPrinter.Properties.label) { $GPOName += " $($GPOPrinter.Properties.label)" }

                $Description = "$($GPOPrinter.desc)".Trim()
                If (!$Description -and $OverrideEmptyDescription) { $Description = $GPOName }

                $GPOSelfHealingEnabled = "1"
                If ($GPOPrinter.Filters.FilterRunOnce -and (!$SelfHealingEnabled)) { $GPOSelfHealingEnabled = "0" }

                $GPOPrinterName = Get-UniqueActionName -ObjectList $VUEMPrinters -ActionName $GPOName

                $VUEMPrinter = New-VUEMPrinterObject -Name "$GPOPrinterName" `
                                                     -Description "$Description" `
                                                     -TargetPath "$($GPOPrinter.Properties.path)" `
                                                     -SelfHealingEnabled "$GPOSelfHealingEnabled" `
                                                     -State "$State" `
                                                     -ObjectList $VUEMPrinters

                If ($VUEMPrinter) { 
                    # add new object to array
                    $VUEMPrinters += $VUEMPrinter
                
                    # grab GPO Filters for this new object
                    ForEach ($Filter in $GPODrive.Filters) {
                        $GPOFilter = New-GPOFilterObject -Name "$GPOPrinterName" -Filter $Filter -ActionType "Printer"
                        If ($GPOFilter) { $GPOFilters += $GPOFilter }
                    }
                }
            }
        }

        # grab Printers from gpreport.xml       
        $GPOPrinters = $GPOReport.GPO.User.ExtensionData.Extension.PrinterConnection.Path

        $GPOSelfHealingEnabled = "1"

        If ($GPOPrinters) {
            Write-Host "Found Deployed Printers in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            ForEach ($GPOPrinter in $GPOPrinters) {
                $GPOPrinterName = Get-UniqueActionName -ObjectList $VUEMPrinters -ActionName "$Prefix$($GPOPrinter.ToString())"
    
                $Description = $null
                If ($OverrideEmptyDescription) { $Description = "$($GPOPrinter.ToString())"}

                $VUEMPrinter = New-VUEMPrinterObject -Name "$GPOPrinterName" `
                                                     -Description "$Description" `
                                                     -TargetPath "$($GPOPrinter.ToString())" `
                                                     -SelfHealingEnabled "$GPOSelfHealingEnabled" `
                                                     -State "$State" `
                                                     -ObjectList $VUEMPrinters
            
                If ($VUEMPrinter) { $VUEMPrinters += $VUEMPrinter }
            }
        }
        #endregion

        #region GPO Preferences - Registry
        If (Test-Path -Path ("{0}Registry\Registry.xml" -f $GPOPreferenceLocation)) {
            Write-Host "Found Registry User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab Registry from Registry.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}Registry\Registry.xml" -f $GPOPreferenceLocation)
            # grab Registry
            $GPORegistrySettings = @()
            If ($GPOPreference.RegistrySettings.Registry) {
                $GPORegistrySettings += $GPOPreference.RegistrySettings.Registry
            }
            # grab and process collections
            If ($GPOPreference.RegistrySettings.Collection) {
                $GPORegistrySettings += Get-GPORegistrySettingsFromCollection -Collections $GPOPreference.RegistrySettings.Collection
            }
            
            # process registry
            ForEach ($GPORegistrySetting in $GPORegistrySettings) {
                If ($GPORegistrySetting.Properties.key -and $GPORegistrySetting.Properties.type -and ($GPORegistrySetting.Properties.type -notlike "REG_BINARY")) {
                    $GPOName = "$Prefix$($GPORegistrySetting.Properties.key)"
                    If ($GPORegistrySetting.Properties.name) { 
                        $GPOName += "\$($GPORegistrySetting.Properties.name)"
                    } Else {
                        $GPOName += "\(Default)"
                    }

                    $GPORegistrySettingRunOnce = "0"
                    If ($GPORegistrySetting.Filters.FilterRunOnce -and (!$SelfHealingEnabled)) { $GPORegistrySettingRunOnce = "1" }

                    $GPORegistrySettingAction = "0"
                    If ($GPORegistrySetting.Properties.action -eq "D") { 
                        $GPORegistrySettingAction = "1"
                        $GPOName += " (Delete)"
                    }

                    $Description = "$($GPORegistrySetting.desc)".Trim()
                    If (!$Description -and $OverrideEmptyDescription) { $Description = "$GPOName" }

                    $GPORegistrySettingName = Get-UniqueActionName -ObjectList $VUEMRegValues -ActionName $GPOName
                    
                    $VUEMRegValue = New-VUEMRegValueObject -Name "$GPORegistrySettingName" `
                                                        -Description "$Description" `
                                                        -TargetName "$($GPORegistrySetting.Properties.name)" `
                                                        -TargetPath "$($GPORegistrySetting.Properties.key)" `
                                                        -TargetType "$($GPORegistrySetting.Properties.type)" `
                                                        -TargetValue "$($GPORegistrySetting.Properties.value)" `
                                                        -RunOnce "$GPORegistrySettingRunOnce" `
                                                        -ActionType "$GPORegistrySettingAction" `
                                                        -State "$State" `
                                                        -ObjectList $VUEMRegValues

                    If ($VUEMRegValue) { 
                        # add new object to array
                        $VUEMRegValues += $VUEMRegValue
                    
                        # grab GPO Filters for this new object
                        ForEach ($Filter in $GPORegistrySetting.Filters) {
                            $GPOFilter = New-GPOFilterObject -Name "$GPORegistrySettingName" -Filter $Filter -ActionType "RegValue"
                            If ($GPOFilter) { $GPOFilters += $GPOFilter }
                        }
                    }
                } elseif ($GPORegistrySetting.Properties.key -and $GPORegistrySetting.Properties.type -and $GPORegistrySetting.Properties.type -like "REG_BINARY") {
                    Write-Verbose "Skipped REG_BINARY Registry Preference: '$($GPORegistrySetting.Properties.name)' - '$($GPORegistrySetting.Properties.key)'"
                }
            }
        }
        #endregion

        #region GPO Preferences - DataSources
        If (Test-Path -Path ("{0}DataSources\DataSources.xml" -f $GPOPreferenceLocation)) {
            Write-Host "Found DataSources User preference xml in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            # grab DataSources from DataSources.xml
            [xml]$GPOPreference = Get-Content -Path ("{0}DataSources\DataSources.xml" -f $GPOPreferenceLocation)
            # grab DataSources where action is not set to D (Delete), driver equals "SQL Server" and User DSN only
            $GPODataSources = $GPOPreference.DataSources.DataSource | Where-Object { $_.Properties.action -notlike "D" -and $_.Properties.driver -eq "SQL Server" }
            
            # convert RegistrySettings to VUEMRegValues
            ForEach ($GPODataSource in $GPODataSources) {
                $GPODataSourceName = Get-UniqueActionName -ObjectList $VUEMNetDrives -ActionName "$Prefix$($GPODataSource.Name)"

                $Description = "$($GPODataSource.Properties.description)".Trim()
                If (!$Description -and $OverrideEmptyDescription) { $Description = "$($GPODataSource.Name)" }

                $GPODataSourceRunOnce = "0"
                If ($GPODataSource.Filters.FilterRunOnce -and (!$SelfHealingEnabled)) { $GPODataSourceRunOnce = "1" }

                $GPODataSourceServerName = ($GPODataSource.Properties.Attributes.Attribute | Where-Object {$_.name -eq "SERVER"}).value
                $GPODataSourceDatabaseName = ($GPODataSource.Properties.Attributes.Attribute | Where-Object {$_.name -eq "DATABASE"}).value

                $VUEMUserDSN = New-VUEMUserDSNObject -Name "$GPODataSourceName" `
                                                     -Description "$Description" `
                                                     -TargetName "$($GPODataSource.Properties.dsn)" `
                                                     -TargetDriverName "$($GPODataSource.Properties.driver)" `
                                                     -TargetServerName "$GPODataSourceServerName" `
                                                     -TargetDatabaseName "$GPODataSourceDatabaseName" `
                                                     -RunOnce "$GPODataSourceRunOnce" `
                                                     -State "$State" `
                                                     -ObjectList $VUEMUserDSNs

                If ($VUEMUserDSN) { 
                    # add new object to array
                    $VUEMUserDSNs += $VUEMUserDSN
                
                    # grab GPO Filters for this new object
                    ForEach ($Filter in $GPODrive.Filters) {
                        $GPOFilter = New-GPOFilterObject -Name "$GPODataSourceName" -Filter $Filter -ActionType "UserDSN"
                        If ($GPOFilter) { $GPOFilters += $GPOFilter }
                    }
                }
            }
        }
        #endregion

        #region GPO - Policy Run these programs at logon
        # grab ProgramsToRun from gpreport.xml
        $GPOPrograms = @()
        $GPORunProgramsAtUserLogon = $GPOReport.GPO.User.ExtensionData.Extension.Policy | Where-Object { $_.Name -like "run these programs at user logon" }
        If ($GPORunProgramsAtUserLogon.ListBox.Value.Element.Data) { $GPOPrograms += $GPORunProgramsAtUserLogon.ListBox.Value.Element.Data }
        $GPORunProgramsAtUserLogon = $GPOReport.GPO.Computer.ExtensionData.Extension.Policy | Where-Object { $_.Name -like "run these programs at user logon" }
        If ($GPORunProgramsAtUserLogon.ListBox.Value.Element.Data) { $GPOPrograms += $GPORunProgramsAtUserLogon.ListBox.Value.Element.Data }

        If ($GPOPrograms) {
            Write-Host "Found 'Run Programs At User Logon' settings in '$($GPOBackup.FullName)'" -ForegroundColor Yellow
            
            ForEach ($GPOProgram in $GPOPrograms) {
                # grab command and args (if any)
                $ProgramTokens = [Management.Automation.PSParser]::Tokenize($GPOProgram, [ref]$null)
                $GPOProgramCommand = $ProgramTokens[0].Content
                $GPOProgramArgs = $GPOProgram.Replace($GPOProgramCommand, "").Trim()
    
                $GPOName = "$Prefix$([System.IO.Path]::GetFileNameWithoutExtension("$($GPOProgramCommand)"))"
                $GPOProgramName = Get-UniqueActionName -ObjectList $VUEMExtTasks -ActionName "$GPOName"
                
                $Description = $null
                If ($OverrideEmptyDescription) { $Description = "$([System.IO.Path]::GetFileNameWithoutExtension("$($GPOProgramCommand)"))" }

                $VUEMExtTask = New-VUEMExtTaskObject -Name "$GPOProgramName" `
                                                     -Description "$Description" `
                                                     -TargetPath "$GPOProgramCommand" `
                                                     -TargetArgs "$GPOProgramArgs" `
                                                     -State "$State" `
                                                     -ObjectList $VUEMExtTasks

                If ($VUEMExtTask) { $VUEMExtTasks += $VUEMExtTask }
            }
        }
        #endregion

        #region GPO - Logon scripts
        # grab LogonScripts from gpreport.xml 
        $GPOScripts = $GPOReport.GPO.User.ExtensionData.Extension.Script | Where-Object { $_.Type -like "logon" }
        
        If ($GPOScripts) {
            Write-Host "Found User Logon Scripts in '$($GPOBackup.FullName)'" -ForegroundColor Yellow

            ForEach ($GPOScript in $GPOScripts) {
                $GPOScriptCommand = $GPOScript.Command
                $GPOScriptArgs = $GPOScript.Parameters.Trim()

                $GPOName = "$Prefix$([System.IO.Path]::GetFileNameWithoutExtension("$($GPOScriptCommand)"))"
                $GPOScriptName = Get-UniqueActionName -ObjectList $VUEMExtTasks -ActionName "$GPOName"

                $Description = $null
                If ($OverrideEmptyDescription) { $Description = "$([System.IO.Path]::GetFileNameWithoutExtension("$($GPOScriptCommand)"))" }
                
                $GPOScriptState = $State
                If (![System.IO.Path]::GetDirectoryName($GPOScriptCommand)) { 
                    $GPOScriptName = "[NEEDS FILE LOCATION]" + $GPOScriptName
                    $GPOScriptState = "0"
                    If (!$GPOScriptFiles.Contains(($GPOBackup.FullName + "\DomainSysvol\GPO\User\Scripts\Logon").ToLower())) { $GPOScriptFiles += ($GPOBackup.FullName + "\DomainSysvol\GPO\User\Scripts\Logon\").ToLower() }
                }

                $VUEMExtTask = New-VUEMExtTaskObject -Name "$GPOScriptName" `
                                                    -Description "$Description" `
                                                    -TargetPath "$GPOScriptCommand" `
                                                    -TargetArgs "$GPOScriptArgs" `
                                                    -State "$GPOScriptState" `
                                                    -ObjectList $VUEMExtTasks
            
                If ($VUEMExtTask) { $VUEMExtTasks += $VUEMExtTask }
            }
        }
        #endregion
    }

    #region output xml files
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
    If ($VUEMExtTasks) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMExtTask" -ObjectList $VUEMExtTasks | Out-File $OutputPath\VUEMExtTasks.xml
        Write-Host "VUEMExtTasks.xml written to '$OutputPath\VUEMExtTasks.xml'" -ForegroundColor Green
    }
    #endregion

    #region Export
    If ($GPOFilters -and $ExportFilters) {
        $GPOFilters | Export-Csv -NoTypeInformation -Path "$OutputPath\GPOFilters.csv"
        Write-Host "GPO Filter data written to '$OutputPath\GPOFilters.csv'" -ForegroundColor Green
    }
    If ($GPOScriptFiles) {
        If (!(Test-Path -Path "$OutputPath\Embedded Logon Script files")) { New-Item -Name "Embedded Logon Script files" -Path $OutputPath -ItemType Directory -Force | Out-Null }

        ForEach ($GPOScriptFile in $GPOScriptFiles) {
            Copy-Item -Path "$GPOScriptFile\*" -Destination "$OutputPath\Embedded Logon Script files\" -Force | Out-Null
        }
        Write-Host "GPO Embedded Logon Script files written to '$OutputPath\Embedded Logon Script files\'" -ForegroundColor Green
    }
    #endregion
}

<#
    .Synopsis
    Imports Environmental Settings from GPOs and converts them to WEM Environmental Settings.

    .Description
    Imports Environmental Settings from GPOs and converts them to WEM Environmental Settings.
    Output will be a VUEMEnvironmentalSettings.xml ready for import in WEM.

    .Link
    https://msfreaks.wordpress.com

    .Parameter GPOBackupPath
    This is the path where the GPO Backup files are stored.
    GPO Backups are each stored in its own folder like {<GPO Backup GUID>}.
    All GPO Backups in the GPOBackupPath are processed, newest to oldest. If duplicate settings are found,
    the oldest are discarded. 

    .Parameter OutputPath
    Location where the output xml file will be written to. Defaults to current folder if
    omitted.

    .Parameter Enable
    If used will create all settings found in the GPO Backups in Enabled state for WEM.
    Use with caution, this will be applied to all agents in the configuration set!

    .Example
    Import-VUEMEnvironmentalSettingsFromGpo -GPOBackupPath C:\GPOBackups

    Description

    -----------

    Create WEM Microsoft USV Settings from all relevant settings found in all the GPOBackups found in C:\GPOBackups.
    Output is created in the current folder.

    .Example
    Import-VUEMEnvironmentalSettingsFromGpo -GPOBackupPath C:\GPOBackups -OutputPath C:\Temp

    Description

    -----------

    Create WEM Environmental Settings from all relevant settings found in all the GPOBackups found in C:\GPOBackups.
    Output is created in c:\temp.

    .Example
    Import-VUEMEnvironmentalSettingsFromGpo -GPOBackupPath C:\GPOBackups -Enable

    Description

    -----------

    Create WEM Environmental Settings from all relevant settings found in all the GPOBackups found in C:\GPOBackups.
    By default the Environmental Settings processing will be disabled but in this case they will be Enabled.
    Use this setting with caution: Importing this XML will automatically enable all imported Environmental Settings
    for all agents in the configuration set.

    .Notes
    Author:  Arjan Mensch
    Version: 1.0.0

    Environmental Settings
    -----------
    This function will process the GPO Backups from newest to oldest.
    If duplicate settings are found, the newest will be used.
    Exclude Administrators is always checked by default.
    Settings are only processed if the GPOs are in English.

    User Interface: Start Menu
    -----------
    Hide Administrative Tools setting is not checked for in the GPO Backups (No such GPO setting).
    Hide Devices and Printers setting is not checked for in the GPO Backups (No such GPO setting).

    User Interface: Appearance
    -----------
    Set Background Color setting is not checked for in the GPO Backups (cannot convert #RGB to WEM Named Colors).

    User Interface: Edge UI
    -----------
    Disable Switcher setting is not checked for in the GPO Backups (No such GPO setting).
    Disable Charms Hint setting is not checked for in the GPO Backups (No such GPO setting).

    User Interface: Explorer
    -----------
    Hide Network icon in Explorer setting is not checked for in the GPO Backups (No such GPO setting).

    User Interface: Control Panel
    -----------
    Hide Control Panel setting cancels out Show only and Hide specified Control Panel applets.
    Show Only specified Control Panel applets cancels out Hide Control Panel and Hide specified Control Panel
    applets.
    Hide specified Control Panel applets cancels out Hide Control Panel and Show only specified Control Panel
    applets.
    GPO Backups are processed for these items in the above order.

    User Interface: Advanced Tuning
    -----------
    SBC / HDV Tuning settings are not checked for in the GPO Backups.
    These are environment specific settings and should be handled in WEM if you want to enable them.
#>
Function Import-VUEMEnvironmentalSettingsFromGpo {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,Position=1)]
        [string]$GPOBackupPath,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$Enable = $False
    )
    Write-Verbose "$(Get-Date -Format G)"
    Write-Verbose "Function: $($MyInvocation.MyCommand)"

    # check if $GPOBackupPath is valid
    If ($GPOBackupPath -and !(Test-Path -Path $GPOBackupPath)) {
        Write-Host "Cannot find path '$GPOBackupPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    }
    if ($GPOBackupPath.EndsWith("\")) { $GPOBackupPath = $GPOBackupPath.Substring(0,$GPOBackupPath.Length-1) }

    # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path -Path $OutputPath) -or ((Get-Item -Path $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    } Elseif ($OutputPath.EndsWith("\")) { $OutputPath = $OutputPath.Substring(0,$OutputPath.Length-1) }

    # grab GPO backups
    $GPOBackups = @()
    $GPOBackups = Get-ChildItem -Path $GPOBackupPath -Directory | Where-Object { $_.Name -like "{*}" } | Sort-Object -Property LastWriteTime -Descending | Select-Object FullName

    If (!$GPOBackups) {
        Write-Host "Cannot locate GPO Backups in '$GPOBackupPath'" -ForegroundColor Red
        Break
    }

    # Verbose output
    Write-Verbose "Using GPOBackupPath '$GPOBackupPath'"
    Write-Verbose "Using OutputPath '$OutputPath'"
    Write-Verbose "Using Enable: $Enable"
    Write-Verbose "Found $($GPOBackups.Count) GPO Backups in '$GPOBackupPath'"

    # init VUEM Configuration Setting array
    $VUEMConfigurationSettings = @()

    #region Init default objects
    # get default objects
    $Value = "0"
    If ($Enable) { $Value = "1" }
    $VUEMConfigurationSettings = Get-VUEMEnvironmentalInitialObjects -Value $Value
    #endregion

    # process all GPO backups in the folder
    Write-Host "Processing GPO Backups in '$GPOBackupPath'" -ForegroundColor Yellow

    ForEach ($GPOBackup in $GPOBackups) {
        Write-Verbose "Processing '$($GPOBackup.FullName)'"

        # get gpreport.xml
        [xml]$GPOReport = ""
        If (Test-Path -Path ("{0}\gpreport.xml" -f $GPOBackup.FullName)) {
            [xml]$GPOReport = Get-Content -Path ("{0}\gpreport.xml" -f $GPOBackup.FullName)
        }

        If ($GPOReport) {
            # Grab all computer, user and folder objects for easier processing
            $GPOComputer = $GPOReport.GPO.Computer.ExtensionData.Extension.Policy
            $GPOUser = $GPOReport.GPO.User.ExtensionData.Extension.Policy

            #region tab Start Menu

            # Set Hide Common Programs
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove common program groups from Start Menu" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "HideCommonPrograms" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove Run menu from Start Menu
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Run menu from Start Menu" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "RemoveRunFromStartMenu" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Hide Administrative Tools
            # See Notes for this function

            # Set Remove Help menu form Start Menu
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Help menu from Start Menu" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "HideHelp" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove Search link from Start Menu
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Search link from Start Menu" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "HideFind" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove links and access to Windows Update
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove links and access to Windows Update" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "HideWindowsUpdate" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Lock the Taskbar
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Lock the Taskbar" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "LockTaskbar" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove Clock from the system notification area
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Clock from the system notification area" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "HideSystemClock" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Hide Devices and Printers
            # See Notes for this function

            # Set Change Start Menu power button
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Change Start Menu power button" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled" -and $GPOPolicy.DropDownList.Value.Name -like "log off") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "HideTurnOff" -Value "1" -ObjectList $VUEMConfigurationSettings
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "ForceLogoff" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Turn off notification area cleanup
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Turn off notification area cleanup" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "Turnoffnotificationareacleanup" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Turn off personalized menus
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Turn off personalized menus" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "TurnOffpersonalizedmenus" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Clear the recent programs list for new users
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Clear the recent programs list for new users" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "ClearRecentprogramslist" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Load a specific theme
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Load a specific theme" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                If ($GPOPolicy.EditText.Value) { $GPOPolicyValue = $GPOPolicy.EditText.Value.Replace("%username%","##username##") }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "SetSpecificThemeFile" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "SpecificThemeFileValue" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Set Force a specific visual style file or force Windows Classic
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Force a specific visual style file or force Windows Classic" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                If ($GPOPolicy.EditText.Value) { $GPOPolicyValue = $GPOPolicy.EditText.Value.Replace("%username%","##username##") }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "SetVisualStyleFile" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "VisualStyleFileValue" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Set Desktop Wallpaper
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Desktop Wallpaper" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                $GPOPolicyStyle = ""
                If ($GPOPolicy.EditText.Value) { $GPOPolicyValue = $GPOPolicy.EditText.Value.Replace("%username%","##username##") }
                If ($GPOPolicy.DropDownList.Value.Name) { 
                    Switch($GPOPolicy.DropDownList.Value.Name.ToLower()) {
                        "fill" {
                            $GPOPolicyStyle = "10"
                            Continue
                        }
                        "center" {
                            $GPOPolicyStyle = "4"
                            Continue
                        }
                        "stretch" {
                            $GPOPolicyStyle = "2"
                            Continue
                        }
                        "fit" {
                            $GPOPolicyStyle = "6"
                            Continue
                        }
                        "tile" {
                            $GPOPolicyStyle = "0"
                            Continue
                        }
                    }
                }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue -and $GPOPolicyStyle) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "SetWallpaper" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "Wallpaper" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "WallpaperStyle" -Value "$GPOPolicyStyle" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Set Force a specific background and accent color
            # See Notes for this function

            #endregion

            #region tab Desktop
            # Set Remove Computer icon on the desktop
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Computer icon on the desktop" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "NoMyComputerIcon" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove Recycle Bin icon from desktop
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Recycle Bin icon from desktop" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "NoRecycleBinIcon" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove My Documents icon on the desktop
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove My Documents icon on the desktop" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "NoMyDocumentsIcon" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Go to the desktop instead of Start when signing in
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Go to the desktop instead of Start when signing in" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "BootToDesktopInsteadOfStart" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove Properties from the Computer icon context menu
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Properties from the Computer icon context menu" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "NoPropertiesMyComputer" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove Properties from the Recycle Bin context menu
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Properties from the Recycle Bin context menu" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "NoPropertiesRecycleBin" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove Properties from the Documents icon context menu
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Properties from the Documents icon context menu" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "NoPropertiesMyDocuments" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove the networking icon
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove the networking icon" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "HideNetworkIcon" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Hide Network Locations icon on desktop
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Hide Network Locations icon on desktop" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "HideNetworkConnections" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove Task Manager
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Task Manager" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "DisableTaskMgr" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Edge UI Disable Switcher

            # Set Edge UI Disable Charms Hint
            
            #endregion

            #region tab Windows Explorer
            # Set Prevent access to registry editing tools / Disable regedit from running silently?
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Prevent access to registry editing tools" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "DisableRegistryEditing" -Value "1" -ObjectList $VUEMConfigurationSettings
            }
            If ($GPOPolicy -and $GPOPolicy.DropDownList.Value.Name -like "yes") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "DisableSilentRegedit" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Prevent access to the command prompt / Disable the command prompt script processing also?
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Prevent access to the command prompt" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "DisableCmd" -Value "1" -ObjectList $VUEMConfigurationSettings
            }
            If ($GPOPolicy -and $GPOPolicy.DropDownList.Value.Name -like "yes") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "DisableCmdScripts" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Hides the Manage item on the File Explorer context menu
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Hides the Manage item on the File Explorer context menu" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "RemoveContextMenuManageItem" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove "Map Network Drive" and "Disconnect Network Drive"
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove ""Map Network Drive"" and ""Disconnect Network Drive""" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "NoNetConnectDisconnect" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Turn off Windows Libraries features that rely on indexed file data
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Turn off Windows Libraries features that rely on indexed file data" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "HideLibrairiesInExplorer" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Hide Network In Explorer
            # See Notes for this function

            # Set Remove Add or Remove Programs
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove Add or Remove Programs" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "NoProgramsCPL" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove Windows Security item from Start Menu
            $GPOPolicy = $GPOComputer | Where-Object { $_.Name -like "Remove Windows Security item from Start Menu" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "NoNtSecurity" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove File Explorer's default context menu
            $GPOPolicy = $GPOuser | Where-Object { $_.Name -like "Remove File Explorer's default context menu" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "NoViewContextMenu" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Remove access to the context menus for the taskbar
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Remove access to the context menus for the taskbar" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "NoTrayContextMenu" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Hide these specified drives in My Computer
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Hide these specified drives in My Computer" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                If ($GPOPolicy.DropDownList.Value.Name) { 
                    Switch ($GPOPolicy.DropDownList.Value.Name) {
                        "Restrict A and B drives only" {
                            $GPOPolicyValue = "A;B"
                            Continue
                        }
                        "Restrict C drive only" {
                            $GPOPolicyValue = "C"
                            Continue
                        }
                        "Restrict D drive only" {
                            $GPOPolicyValue = "D"
                            Continue
                        }
                        "Restrict A, B, C drives only" {
                            $GPOPolicyValue = "A;B;C"
                            Continue
                        }
                        "Restrict A, B, C and D drives only" {
                            $GPOPolicyValue = "A;B;C;D"
                            Continue
                        }
                        "Restrict all drives" {
                            $GPOPolicyValue = "A;B;C;D;E;F;G;H;I;J;K;L;M;N;O;P;Q;R;S;T;U;V;W;X;Y;Z"
                            Continue
                        }
                    }
                }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "HideSpecifiedDrivesFromExplorer" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "ExplorerHiddenDrives" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Set Prevent access to drives from My Computer
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Prevent access to drives from My Computer" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                If ($GPOPolicy.DropDownList.Value.Name) { 
                    Switch ($GPOPolicy.DropDownList.Value.Name) {
                        "Restrict A and B drives only" {
                            $GPOPolicyValue = "A;B"
                            Continue
                        }
                        "Restrict C drive only" {
                            $GPOPolicyValue = "C"
                            Continue
                        }
                        "Restrict D drive only" {
                            $GPOPolicyValue = "D"
                            Continue
                        }
                        "Restrict A, B, C drives only" {
                            $GPOPolicyValue = "A;B;C"
                            Continue
                        }
                        "Restrict A, B, C and D drives only" {
                            $GPOPolicyValue = "A;B;C;D"
                            Continue
                        }
                        "Restrict all drives" {
                            $GPOPolicyValue = "A;B;C;D;E;F;G;H;I;J;K;L;M;N;O;P;Q;R;S;T;U;V;W;X;Y;Z"
                            Continue
                        }
                    }
                }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "RestrictSpecifiedDrivesFromExplorer" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "ExplorerRestrictedDrives" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                }
            }

            #endregion

            #region tab Control Panel
            # Set Prohibit access to Control Panel and PC settings
            # See Notes for this function
            $GPOPolicy = $GPOuser | Where-Object { $_.Name -like "Prohibit access to Control Panel and PC settings" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "HideControlPanel" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Set Show only specified Control Panel items
            # See Notes for this function
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Show only specified Control Panel items" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                If ($GPOPolicy.Listbox.Value.Element.Data) { $GPOPolicyValue = $GPOPolicy.Listbox.Value.Element.Data -join ";" }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "RestrictCpl" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "RestrictCplList" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Set Hide specified Control Panel items
            # See Notes for this function
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Hide specified Control Panel items" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                If ($GPOPolicy.Listbox.Value.Element.Data) { $GPOPolicyValue = $GPOPolicy.Listbox.Value.Element.Data -join ";" }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "DisallowCpl" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "0" -Name "DisallowCplList" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                }
            }

            #endregion

            #region tab Known Folder Management
            # Set Disable Known Folders
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Disable Known Folders" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                If ($GPOPolicy.Listbox.Value.Element.Data) { $GPOPolicyValue = $GPOPolicy.Listbox.Value.Element.Data -join ";" }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "DisableSpecifiedKnownFolders" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "DisabledKnownFolders" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                }
            }

            #endregion

            #region tab SBC / HDV Tuning

            # See Notes for this function

            #endregion
        }
    }
    If ($VUEMConfigurationSettings | Where-Object { $_.Init -eq $False }) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMEnvironmentalSetting" -ObjectList $VUEMConfigurationSettings | Out-File $OutputPath\VUEMEnvironmentalSettings.xml
        Write-Host "VUEMEnvironmentalSettings.xml written to '$OutputPath\VUEMEnvironmentalSettings.xml'" -ForegroundColor Green
    }

}

<# 
    .Synopsis
    Imports Microsoft User State Virtualization settings from GPOs and converts them to WEM Microsoft User State
    Virtualization Settings.

    .Description
    Imports Microsoft User State Virtualization settings from GPOs and converts them to WEM Microsoft User State
    Virtualization Settings.
    Output will be a VUEMMicrosftUsvSettings.xml ready for import in WEM.

    .Link
    https://msfreaks.wordpress.com

    .Parameter GPOBackupPath
    This is the path where the GPO Backup files are stored.
    GPO Backups are each stored in its own folder like {<GPO Backup GUID>}.
    All GPO Backups in the GPOBackupPath are processed, newest to oldest. If duplicate settings are found,
    the oldest are discarded. 

    .Parameter OutputPath
    Location where the output xml file will be written to. Defaults to current folder if
    omitted.

    .Parameter Enable
    If used will create all settings found in the GPO Backups in Enabled state for WEM.
    Use with caution, this will be applied to all agents in the configuration set!

    .Example
    Import-VUEMMicrosoftUsvSettingsFromGpo -GPOBackupPath C:\GPOBackups

    Description

    -----------

    Create WEM Microsoft USV Settings from all relevant settings found in all the GPOBackups found in C:\GPOBackups.
    Output is created in the current folder.

    .Example
    Import-VUEMMicrosoftUsvSettingsFromGpo -GPOBackupPath C:\GPOBackups -OutputPath C:\Temp

    Description

    -----------

    Create WEM Microsoft USV Settings from all relevant settings found in all the GPOBackups found in C:\GPOBackups.
    Output is created in c:\temp.

    .Example
    Import-VUEMMicrosoftUsvSettingsFromGpo -GPOBackupPath C:\GPOBackups -Enable

    Description

    -----------

    Create WEM Microsoft USV Settings from all relevant settings found in all the GPOBackups found in C:\GPOBackups.
    By default the Microsoft USV Settings processing will be disabled but in this case they will be Enabled.
    Use this setting with caution: Importing this XML will automatically enable all imported Microsoft USV Settings
    for all agents in the configuration set.

    .Notes
    Author:  Arjan Mensch
    Version: 1.0.0

    Microsoft User State Virtualization settings
    -----------
    This function will process the GPO Backups from newest to oldest.
    If duplicate settings are found, the newest will be used.
    Exclude Administrators is always checked by default.
    Settings are only processed if the GPOs are in English.

    Folder Redirection
    -----------
    Delete local Redirected Folders setting is not checked for in the GPO Backups (No such GPO setting).
#>
Function Import-VUEMMicrosoftUsvSettingsFromGpo {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,Position=1)]
        [string]$GPOBackupPath,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$Enable = $False
    )
    Write-Verbose "$(Get-Date -Format G)"
    Write-Verbose "Function: $($MyInvocation.MyCommand)"

    # check if $GPOBackupPath is valid
    If ($GPOBackupPath -and !(Test-Path -Path $GPOBackupPath)) {
        Write-Host "Cannot find path '$GPOBackupPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    }
    if ($GPOBackupPath.EndsWith("\")) { $GPOBackupPath = $GPOBackupPath.Substring(0,$GPOBackupPath.Length-1) }

    # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path -Path $OutputPath) -or ((Get-Item -Path $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    } Elseif ($OutputPath.EndsWith("\")) { $OutputPath = $OutputPath.Substring(0,$OutputPath.Length-1) }

    # grab GPO backups
    $GPOBackups = @()
    $GPOBackups = Get-ChildItem -Path $GPOBackupPath -Directory | Where-Object { $_.Name -like "{*}" } | Sort-Object -Property LastWriteTime -Descending | Select-Object FullName

    If (!$GPOBackups) {
        Write-Host "Cannot locate GPO Backups in '$GPOBackupPath'" -ForegroundColor Red
        Break
    }

    # Verbose output
    Write-Verbose "Using GPOBackupPath '$GPOBackupPath'"
    Write-Verbose "Using OutputPath '$OutputPath'"
    Write-Verbose "Using Enable: $Enable"
    Write-Verbose "Found $($GPOBackups.Count) GPO Backups in '$GPOBackupPath'"
    
    # init VUEM Configuration Setting array
    $VUEMConfigurationSettings = @()

    #region Init default objects
    # get default objects
    $Value = "0"
    If ($Enable) { $Value = "1" }
    $VUEMConfigurationSettings = Get-VUEMMicrosoftUsvInitialObjects -Value $Value
    #endregion

    # process all GPO backups in the folder
    Write-Host "Processing GPO Backups in '$GPOBackupPath'" -ForegroundColor Yellow

    ForEach ($GPOBackup in $GPOBackups) {
        Write-Verbose "Processing '$($GPOBackup.FullName)'"

        # get gpreport.xml
        [xml]$GPOReport = ""
        If (Test-Path -Path ("{0}\gpreport.xml" -f $GPOBackup.FullName)) {
            [xml]$GPOReport = Get-Content -Path ("{0}\gpreport.xml" -f $GPOBackup.FullName)
        }

        If ($GPOReport) {
            # Grab all computer, user and folder objects for easier processing
            $GPOComputer = $GPOReport.GPO.Computer.ExtensionData.Extension.Policy
            $GPOUser = $GPOReport.GPO.User.ExtensionData.Extension.Policy
            $GPOFolders = $GPOReport.GPO.User.ExtensionData.Extension.Folder

            # variable to later check if any folders were redirected
            $GPOFolderRedirection = $False

            #region tab Roaming Profiles Configuration

            # Set Windows Roaming Profiles Path
            $GPOPolicy = $GPOComputer | Where-Object { $_.Name -like "Set roaming profile path for all users logging onto this computer" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                If ($GPOPolicy.EditText.Value) { $GPOPolicyValue = $GPOPolicy.EditText.Value.Replace("%username%","##username##") }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "SetWindowsRoamingProfilesPath" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "WindowsRoamingProfilesPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Set RDS Roaming Profiles Path
            $GPOPolicy = $GPOComputer | Where-Object { $_.Name -like "Set path for Remote Desktop Services Roaming User Profile" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                If ($GPOPolicy.EditText.Value) { $GPOPolicyValue = $GPOPolicy.EditText.Value }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "SetRDSRoamingProfilesPath" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "RDSRoamingProfilesPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Set RDS Home Drive Path
            $GPOPolicy = $GPOComputer | Where-Object { $_.Name -like "Set Remote Desktop Services User Home Directory" }
            If ($GPOPolicy) {
                $GPOPolicyDriveLetter = ""
                $GPOPolicyValue = ""
                If ($GPOPolicy.EditText.Value) { $GPOPolicyValue = $GPOPolicy.EditText.Value }
                If (($GPOPolicy.DropDownList | Where-Object {$_.Name -like "Drive Letter" })) { $GPOPolicyDriveLetter = ($GPOPolicy.DropDownList | Where-Object {$_.Name -like "Drive Letter" }).Value.Name }

                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue -and $GPOPolicyDriveLetter) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "SetRDSHomeDrivePath" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "RDSHomeDrivePath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "RDSHomeDriveLetter" -Value "$GPOPolicyDriveLetter" -ObjectList $VUEMConfigurationSettings
                }
            }
            #endregion

            #region tab Roaming Profiles Advanced Configuration

            # Enable Folders Exclusions
            $GPOPolicy = $GPOUser | Where-Object { $_.Name -like "Exclude directories in roaming profile" }
            If ($GPOPolicy) {
                $GPOPolicyValue = ""
                If ($GPOPolicy.EditText.Value) { $GPOPolicyValue = $GPOPolicy.EditText.Value.Replace("%username%","##username##") }
                If ($GPOPolicy.State -like "enabled" -and $GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "2" -Name "SetRoamingProfilesFoldersExclusions" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "2" -Name "RoamingProfilesFoldersExclusions" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Delete Cached Copies Of Roaming Profiles
            $GPOPolicy = $GPOComputer | Where-Object { $_.Name -like "Delete cached copies of roaming profiles" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") { 
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "DeleteRoamingCachedProfiles" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Add the Administrators Security Group to Roaming User Profiles
            $GPOPolicy = $GPOComputer | Where-Object { $_.Name -like "Add the Administrators security group to roaming user profiles" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "AddAdminGroupToRUP" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Do Not Check for User Ownership Of Roaming Profile Folders
            $GPOPolicy = $GPOComputer | Where-Object { $_.Name -like "Do not check for user ownership of Roaming Profile Folders" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "CompatibleRUPSecurity" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Do Not Detect Slow Network Connections
            $GPOPolicy = $GPOComputer | Where-Object { $_.Name -like "Disable detection of slow network connections" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "DisableSlowLinkDetect" -Value "1" -ObjectList $VUEMConfigurationSettings
            }

            # Wait for Remote User Profile
            $GPOPolicy = $GPOComputer | Where-Object { $_.Name -like "Wait for remote user profile" }
            If ($GPOPolicy -and $GPOPolicy.State -like "enabled") {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "1" -Name "SlowLinkProfileDefault" -Value "1" -ObjectList $VUEMConfigurationSettings
            }
            #endregion

            #region tab Folder Redirection

            # Redirect Desktop
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath) { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processDesktopRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "DesktopRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
            }

            # Redirect Start Menu
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath) { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processStartMenuRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "StartMenuRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
            }

            # Redirect Documents
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath) { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processPersonalRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "PersonalRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
            }

            # Redirect Pictures
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{33E28130-4E1E-4676-835A-98395C3BC3BB}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath -and $GPOFolder.Location.DestinationPath -notlike "my pictures") { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue -or $GPOFolder.FollowParent -like "true") {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processPicturesRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "PicturesRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
                If (!$GPOPolicyValue -and $GPOFolder.FollowParent -like "true") {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "MyPicturesFollowsDocuments" -Value "1" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Redirect Music
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{4BD8D571-6D19-48D3-BE97-422220080E43}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath -and $GPOFolder.Location.DestinationPath -notlike "my music") { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue -or $GPOFolder.FollowParent -like "true") {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processMusicRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "MusicRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
                If (!$GPOPolicyValue -and $GPOFolder.FollowParent -like "true") {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "MyMusicFollowsDocuments" -Value "1" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Redirect Videos
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath -and $GPOFolder.Location.DestinationPath -notlike "my videos") { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue -or $GPOFolder.FollowParent -like "true") {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processVideoRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "VideoRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
                If (!$GPOPolicyValue -and $GPOFolder.FollowParent -like "true") {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "MyVideoFollowsDocuments" -Value "1" -ObjectList $VUEMConfigurationSettings
                }
            }

            # Redirect Favorites
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{1777F761-68AD-4D8A-87BD-30B759FA33DD}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath) { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processFavoritesRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "FavoritesRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
            }

            # Redirect AppData (Roaming)
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath) { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processAppDataRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "AppDataRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
            }

            # Redirect Contacts
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{56784854-C6CB-462B-8169-88E350ACB882}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath) { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processContactsRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "ContactsRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
            }

            # Redirect Downloads
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{374DE290-123F-4565-9164-39C4925E467B}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath) { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processDownloadsRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "DownloadsRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
            }

            # Redirect Links
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath) { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processLinksRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "LinksRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
            }

            # Redirect Searches
            $GPOFolder = $GPOFolders | Where-Object { $_.Id -like "{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}" }
            If ($GPOFolder) {
                $GPOPolicyValue = ""
                If ($GPOFolder.Location.DestinationPath) { $GPOPolicyValue = $GPOFolder.Location.DestinationPath.ToLower().Replace("%username%","##username##") }

                If ($GPOPolicyValue) {
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processSearchesRedirection" -Value "1" -ObjectList $VUEMConfigurationSettings
                    $VUEMConfigurationSettings = `
                    New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "SearchesRedirectedPath" -Value "$GPOPolicyValue" -ObjectList $VUEMConfigurationSettings

                    $GPOFolderRedirection = $True
                }
            }

            # set default setting for folderredirection
            If ($GPOFolderRedirection) {
                $VUEMConfigurationSettings = `
                New-VUEMConfigurationSettingObject -State "1" -Type "3" -Name "processFoldersRedirectionConfiguration" -Value "1" -ObjectList $VUEMConfigurationSettings
            }
            #endregion

        }
    }
    If ($VUEMConfigurationSettings | Where-Object { $_.Init -eq $False }) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMUSVConfigurationSetting" -ObjectList $VUEMConfigurationSettings | Out-File $OutputPath\VUEMMicrosoftUsvSettings.xml
        Write-Host "VUEMMicrosoftUsvSettings.xml written to '$OutputPath\VUEMMicrosoftUsvSettings.xml'" -ForegroundColor Green
    }
}

<# 
    .Synopsis
    Builds an .xml file containing WEM Action definitions.

    .Description
    Builds an .xml file containing WEM Action definitions for application shortcuts.
    This function supports multiple types of input and creates the xml file containing
    the Actions ready for import into WEM.

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
    Whether the script needs to recursively process $Path. Only valid when the Path parameter
    is a folder. Defaults to $True if omitted.

    .Parameter FileTypes
    Provide a comma separated list of filetypes to process. Only valid when the Path
    parameter is a folder. If omitted .lnk will be used by default.

    .Parameter Prefix
    Provide a prefix string used to generate Action names (as displayed in the WEM console).

    .Parameter SelfHealingEnabled
    If used will create Actions using the SelfHealingEnabled parameter. Defaults to $False if omitted.

    .Parameter OverrideEmptyDescription
    If used will generate a description based on the Action name, but only if a description is not found
    during processing. Defaults to $False if omitted.

    .Parameter Disable
    If used will create disabled Actions. Defaults to $False if omitted (create Enabled Actions).

    .Example
    New-VUEMApplicationsXml

    Description

    -----------

    Create WEM Actions from all the items in the default Start Menu locations and export this to VUEMApplications.xml
    in the current folder.

    .Example
    New-VUEMApplicationsXml -Prefix "ITW - "

    Description

    -----------

    Create WEM Actions from all the items in the default Start Menu locations and export this to VUEMApplications.xml
    in the current folder.
    All Action names in the WEM Console are prefixed with "ITW - ", like for example "ITW - Notepad".

    .Example
    New-VUEMApplicationsXml -Path "E:\Custom Folder\Menu Items" -FileTypes exe,lnk

    Description

    -----------

    Create VUEMApplications.xml in the current folder from all the items in a custom folder, processing .exe
    and .lnk files.

    .Example
    New-VUEMApplicationsXml -Path "C:\Windows\System32\notepad.exe" -Name "Notepad example"

    Description

    -----------

    Create a WEM Action for Notepad.exe

    .Example
    New-VUEMApplicationsXml -OutputPath "C:\Temp" -OutputFileName "applications.xml"

    Description

    -----------

    Create applications.xml in c:\temp for all the items in the default Start Menu locations.

    .Example
    New-VUEMApplicationsXml -SelfHealingEnabled -Disable

    Description

    -----------

    Create WEM Actions from all the items in the default Start Menu locations and export this to VUEMApplications.xml
    in the current folder.
    Actions are created using the Enable SelfHealing switch, and are disabled.

    .Example
    New-VUEMApplicationsXml -OverrideEmptyDescription

    Description

    -----------

    Create WEM Actions from all the items in the default Start Menu locations and export this to VUEMApplications.xml
    in the current folder.
    If no Description was found during processing, the Action name is used as description.

    .Notes
    Author:  Arjan Mensch
    Version: 1.0.0

    By default, if no Path is given the default Start Menu locations will be processed.
    If Folder Redirection for the Start Menu folder is detected, that folder will be used instead.
#>
Function New-VUEMApplicationsXml {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False,ValueFromPipeline=$True,Position=1)]
        [string]$Path = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputFileName = "VUEMApplications.xml",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [bool]$Recurse = $True,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string[]]$FileTypes = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$Prefix = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$SelfHealingEnabled = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$OverrideEmptyDescription = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$Disable = $False
    )
    Write-Verbose "$(Get-Date -Format G)"
    Write-Verbose "Function: $($MyInvocation.MyCommand)"
    
    # check if $Path is valid
    If ($Path -and !(Test-Path -Path $Path)) {
         Write-Host "Cannot find path '$Path' because it does not exist or is not a valid path." -ForegroundColor Red
         Break
    }

    # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path -Path $OutputPath) -or ((Get-Item -Path $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    } Elseif ($OutputPath.EndsWith("\")) { $OutputPath = $OutputPath.Substring(0,$OutputPath.Length-1) }

    # check if $OutputFileName is valid
    If ($OutputFileName -and (!($OutputFileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -eq -1))) {
        Write-Host "'$OutputFileName' is not a valid filename." -ForegroundColor Red
        Break
    }

    # Verbose output
    Write-Verbose "Using Path '$Path'"
    Write-Verbose "Using OutputPath '$OutputPath'"
    Write-Verbose "Using OutputFileName: '$OutputFileName'"
    Write-Verbose "Using Recurse $Recurse"
    Write-Verbose "Using FileTypes '$FileTypes'"
    Write-Verbose "Using Prefix '$Prefix'"
    Write-Verbose "Using SelfHealingEnabled: $SelfHealingEnabled"
    Write-Verbose "Using OverrideEmptyDescription: $OverrideEmptyDescription"
    Write-Verbose "Using Disable: $Disable"
    
    # grab files
    $files = @()
    
    # tell user we're processing the Start Menu if $Path was not provided
    If (!$Path) {
        If ([System.Environment]::GetFolderPath('StartMenu') -eq "$($env:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Start Menu") {
            Write-Host "`nProcessing default Start Menu folders" -ForegroundColor Yellow
            $files = Get-FilesToProcess("$($env:ProgramData)\Microsoft\Windows\Start Menu\Programs")
            $files += Get-FilesToProcess("$($env:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs")
        } Else {
            Write-Host "`nProcessing redirected Start Menu ($([System.Environment]::GetFolderPath('StartMenu')))" -ForegroundColor Yellow
            $files = Get-FilesToProcess("$([System.Environment]::GetFolderPath('StartMenu'))\Programs")
        }
    } Else {
        Write-Host "`nProcessing '$Path'" -ForegroundColor Yellow
        $files = Get-FilesToProcess($Path)
    }
    Write-Verbose "Found $($Files.Count) files to process"

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
            If (!$Description -and $OverrideEmptyDescription) { $Description = "$($file.BaseName)" }

            $VUEMApplication = New-VUEMApplicationObject -Name "$VUEMAppName" `
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
                                                         -State "$State" `
                                                         -ObjectList $VUEMApplications
    
            If ($VUEMApplication ) { $VUEMApplications += $VUEMApplication }
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
    Imports an exported BrokerApplication CSV and converts this to WEM Applications Actions.

    .Description
    Imports an exported BrokerApplication CSV and converts this to WEM Applications Actions.

    .Link
    https://msfreaks.wordpress.com

    .Parameter CSVFile
    This is the full path including the filename to an exported BrokerApplications CSV file.

    .Parameter OutputPath
    Location where the output xml file will be written to. Defaults to current folder if
    omitted.

    .Parameter OutputFileName
    The default filename is VUEMApplications.xml. Use this parameter to override this if needed. 

    .Parameter Prefix
    Provide a prefix string used to generate Action names (as displayed in the WEM console).

    .Parameter SelfHealingEnabled
    This will enable the SelfHealing option.
    
    .Parameter OverrideEmptyDescription
    If used will generate a description based on the Action name, but only if a description is not found during
    processing.

    .Parameter OverrideDisplayName
    Use this parameter to select a different column in the CSV to provide the DisplayName in the Application Action.
    Possible values here are "Name", "BrowserName" or "ApplicationName". Will use "PublishedName" if omitted.

    .Parameter IgnoreStartMenuFolder
    Use this parameter if you wish to ignore the StartMenuFolder in the CSV file.
    This will create all Application Actions in the default location, which is in the root of the StartMenu.

    .Parameter Disable
    If used will create disabled Actions. Defaults to $False if omitted (uses Enabled status from the imported CSV).

    .Example
    Import-VUEMActionsFromBrokerApplicationCSV -CSVFile C:\BrokerApplications\Export.csv

    Description

    -----------

    Create WEM Applications Actions from the CSV file "Export.csv" in C:\BrokerApplications.
    Output is created in the current folder.

    .Example
    Import-VUEMActionsFromBrokerApplicationCSV -CSVFile C:\BrokerApplications\Export.csv -OutputPath "C:\Temp"

    Description

    -----------

    Create WEM Applications Actions from the CSV file "Export.csv" in C:\BrokerApplications.
    Output is created in c:\temp folder.

    .Example
    Import-VUEMActionsFromBrokerApplicationCSV -CSVFile C:\BrokerApplications\Export.csv -OutputPath "C:\Temp" -OutputFileName "applications.xml"

    Description

    -----------

    Create WEM Applications Actions from the CSV file "Export.csv" in C:\BrokerApplications.
    Output is created in c:\temp folder and will be named applications.xml instead of vuemapplications.xml (which is the default and only filename
    you can use to import!).

    .Example
    Import-VUEMActionsFromBrokerApplicationCSV -CSVFile C:\BrokerApplications\Export.csv -Prefix "ITW - " -Disable

    Description

    -----------

    Create WEM Applications Actions from the CSV file "Export.csv" in C:\BrokerApplications.
    Output is created in the current folder.
    All Action names in the WEM Console are prefixed with "ITW - ", like for example "ITW - Microoft Outlook 2016", and all
    actions are disabled.

    .Example
    Import-VUEMActionsFromBrokerApplicationCSV -CSVFile C:\BrokerApplications\Export.csv -SelfHealingEnabled -OverrideEmptyDescription

    Description

    -----------

    Create WEM Applications Actions from the CSV file "Export.csv" in C:\BrokerApplications.
    Output is created in the current folder.
    All Actions will have SelfHealing set to enabled.
    If no Description was found during processing, the Action name is used as description.

    .Notes
    Author:  Arjan Mensch
    Version: 1.1.0

    Create a valid export file for all Published Applications:
    Get-BrokerApplication [-AdminAddress <remote hostname>] | Export-Csv <filename including path for CSV>

    Refer to the Citrix documentation for Get-BrokerApplication for more info on filtering your output:
    https://citrix.github.io/delivery-controller-sdk/Broker/Get-BrokerApplication/
    
#>
Function Import-VUEMActionsFromBrokerApplicationCSV {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,Position=1)]
        [string]$CSVFile,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputFileName = "VUEMApplications.xml",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$Prefix = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$SelfHealingEnabled = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$OverrideEmptyDescription = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [ValidateSet("Name", "BrowserName", "ApplicationName")]
        [string]$OverrideDisplayName = "PublishedName",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$IgnoreStartMenuFolder = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$Disable = $False
    )
    Write-Verbose "$(Get-Date -Format G)"
    Write-Verbose "Function: $($MyInvocation.MyCommand)"

    # check if $CSVFile is valid
    If ($CSVFile -and (!(Test-Path -Path $CSVFile) -or !($CSVFile.EndsWith(".csv")))) {
        Write-Host "Cannot find file '$CSVFile' because it does not exist or is not a valid CSV file." -ForegroundColor Red
        Break
    }
    
    # set ImportPath
    $ImportPath = Split-Path "$($CSVFile)"

    # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path -Path $OutputPath) -or ((Get-Item -Path $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    } Elseif ($OutputPath.EndsWith("\")) { $OutputPath = $OutputPath.Substring(0,$OutputPath.Length-1) }

    # check if $OutputFileName is valid
    If ($OutputFileName -and (!($OutputFileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -eq -1))) {
        Write-Host "'$OutputFileName' is not a valid filename." -ForegroundColor Red
        Break
    }

    # grab BrokerApplications
    $BrokerApplications = @()
    $BrokerApplications = Import-Csv -Path $CSVFile

    # Verbose output
    Write-Verbose "Using CSVFile '$CSVFile'"
    Write-Verbose "Using OutputPath '$OutputPath'"
    Write-Verbose "Using OutputFileName: '$OutputFileName'"
    Write-Verbose "Using Prefix '$Prefix'"
    Write-Verbose "Using SelfHealingEnabled: $SelfHealingEnabled"
    Write-Verbose "Using OverrideEmptyDescription: $OverrideEmptyDescription"
    Write-Verbose "Using OverrideDisplayName: $OverrideDisplayName"
    Write-Verbose "Using IgnoreStartMenuFolder: $IgnoreStartMenuFolder"
    Write-Verbose "Using Disable: $Disable"
    Write-Verbose "Found $($BrokerApplications.Count) BrokerApplications in '$CSVFile'"

    If (!$BrokerApplications) {
        Write-Host "Cannot locate BrokerApplications in '$CSVFile'" -ForegroundColor Red
        Break
    }

    # check if all the required columns are in the CSV file
    $ColumnsExpected = @("Name", "BrowserName", "ApplicationName", "PublishedName", "ApplicationType", "CommandLineExecutable", "CommandLineArguments", "WorkingDirectory", "Enabled", "StartMenuFolder")
    $ColumnsOK = $True
    $ColumnsCsv = $BrokerApplications | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    $ColumnsExpected | ForEach-Object {
        If ($ColumnsCsv -notcontains $_) {
            $ColumnsOK = $False
            Write-Host "Expected column not found in '$CSVFile': '$($_)'" -ForegroundColor Red
        }
    }
    If (-not $ColumnsOK) {
        Break
    }

    # pre-load System.Drawing namespace
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $IconStreamGeneric = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAEaSURBVFhH7ZTbCoJAEIaFCCKCCKJnLTpQVBdB14HQ00T0CqUP4AN41puJAVe92F3HRZegHfgQFvH7/1nQMmPmZ+Z8uYJOCm01vJe64PF8cZ+Ftho89DxPC8IAeZ73QpZlJWmattsAfsBavsk0yRsD3Ox7ST3A4uTC/OjC7ODCdO/AZOfAeOvAaPOB4foDg1UVwLZtIUmSqG2AIq9vgNcc5coBKHIWgNec0RhAdAUUOSJrjsRxrLYBihxBMa85QzkARY7ImjOkAURXQJEjKOY1Z0RRpLYBihyRNUe5cgCKHEEprzmjMYDoCqjImiNhGKptgApvA3V57wFkzbUGEMmDIGgfAKH84ShypQBdyn3fFwfQSaE1Y+bvx7K+efsbU5+Ow3MAAAAASUVORK5CYII="

    # init VUEM action arrays
    $VUEMApplications = @()

    # define selfhealing
    $SelfHeal = "0"
    If ($SelfHealingEnabled) { $SelfHeal = "1" }

    # process inputs
    ForEach ($app in $BrokerApplications) {
        # define state
        $State = "1"
        If ($Disable -or $app.Enabled -like "False") { $State = "0" }

        # define Description
        $Description = $app.Description
        If ($Description -like "KEYWORDS:*") { $Description = "" }

        # define DisplayName
        $DisplayName = ""
        switch ($OverrideDisplayName) {
            "ApplicationName" { $DisplayName = $app.ApplicationName }
            "BrowserName" { $DisplayName = $app.BrowserName }
            "Name" { $DisplayName = $app.Name }
            "PublishedName" { $DisplayName = $app.PublishedName }
            Default {}
        }

        # define StartMenuFolder
        $StartMenuFolder = ""
        If (!$IgnoreStartMenuFolder -and $app.StartMenuFolder) { $StartMenuFolder = "\$($app.StartMenuFolder)" }

        # define Application parameters
        $Arguments = ""
        $TargetFileName = ""
        $IconLocation = ""
        $IconStream = ""

        $TargetPath = [System.Environment]::ExpandEnvironmentVariables($app.CommandLineExecutable)
        $HotKey = "None"

        $WorkingDirectory = "Url"
        # all set for published content. grab more properties for hosted apps
        If ($app.ApplicationType -like "HostedOnDesktop") {
            $TargetFileName = Split-Path -Path $TargetPath -Leaf 
            $Arguments = $app.CommandLineArguments
            $WorkingDirectory = [System.Environment]::ExpandEnvironmentVariables($app.WorkingDirectory)
            # define IconLocation
            $IconLocation = $TargetPath
            If (Test-Path -Path "$($ImportPath)\$($TargetFileName.SubString(0,$TargetFileName.LastIndexOf("."))).ico") { $IconLocation = "$($ImportPath)\$($TargetFileName.SubString(0,$TargetFileName.LastIndexOf("."))).ico" }
        }

        # only work if we have a target
        If ($TargetPath) {
            # grab icon
            If ($IconLocation -and (Test-Path $IconLocation)) { 
                $IconStream = Get-IconToBase64 ([System.Drawing.Icon]::ExtractAssociatedIcon("$($IconLocation)"))
            } else {
                $IconLocation = "C:\PlaceHolderUsedBecauseNoIconWasFound.exe" 
                $IconStream = $IconStreamGeneric
            }
            
            $VUEMAppName = Get-UniqueActionName -ObjectList $VUEMApplications -ActionName "$Prefix$($app.Name)"
            If ((!$Description) -and $OverrideEmptyDescription) { $Description = "$($app.Name)" }

            $VUEMApplication = New-VUEMApplicationObject -Name "$VUEMAppName" `
                                                            -Description "$Description" `
                                                            -DisplayName "$DisplayName" `
                                                            -StartMenuTarget "Start Menu\Programs$($StartMenuFolder)" `
                                                            -TargetPath "$TargetPath" `
                                                            -Parameters "$Arguments" `
                                                            -WorkingDirectory "$WorkingDirectory" `
                                                            -Hotkey "$Hotkey" `
                                                            -IconLocation "$IconLocation" `
                                                            -IconStream "$IconStream" `
                                                            -SelfHealingEnabled "$SelfHeal" `
                                                            -State "$State" `
                                                            -ObjectList $VUEMApplications
    
            If ($VUEMApplication ) { $VUEMApplications += $VUEMApplication }
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
    Will only process the Mapped Drive letter defined by the DriveLetter value.

    .Parameter OutputPath
    Location where the output xml file will be written to. Defaults to current folder if
    omitted.

    .Parameter OutputFileName
    The default filename is VUEMNetDrives.xml. Use this parameter to override this if needed. 

    .Parameter InputCsv
    A csv file containing at least one field: TargetPath (Mandatory). Field DisplayName is optional.
    For each TargetPath a NetDrive object will be created, with DisplayName, if provided, as DisplayName
    and written to the output xml file.
    Parameter DriveLetter will be ignored if this parameter is used.

    .Parameter Prefix
    Provide a prefix string used to generate Action names (as displayed in the WEM console).

    .Parameter SelfHealingEnabled
    If used will create Actions using the SelfHealingEnabled parameter. Defaults to $False if omitted.

    .Parameter OverrideEmptyDescription
    If used will generate a description based on the Action name, but only if a description is not found
    during processing.

    .Parameter Disable
    If used will create disabled Actions. Defaults to $False if omitted (create Enabled Actions).

    .Example
    New-VUEMNetDrivesXml

    Description

    -----------

    Create WEM Actions from all the Drive Mappings for the current user and export this to VUEMNetDrives.xml in
    the current folder.

    .Example
    New-VUEMNetDrivesXml -Prefix "ITW - "

    Description

    -----------

    Create WEM Actions from all the Drive Mappings for the current user and export this to VUEMNetDrives.xml in
    the current folder.
    All Action names in the WEM Console are prefixed with "ITW - ", like for example "ITW - \\FileServer\Citrix.Wem".

    .Example
    New-VUEMNetDrivesXml -DriveLetter "P"

    Description

    -----------

    Create a WEM Action for only the P: Drive Mapping for the current user and export this to VUEMNetDrives.xml
    in the current folder.

    .Example
    New-VUEMNetDrivesXml -OutputPath "C:\Temp" -OutputFileName "drives.xml"

    Description

    -----------

    Create drives.xml in c:\temp containing WEM Actions for all the Drive Mappings for the current user.

    .Example
    New-VUEMNetDrivesXml -InputCSV "C:\Temp\mappeddrives.csv"

    Description

    -----------

    Create WEM Actions from all the items in C:\Temp\mappeddrives.csv and export these as NetDrive objects to
    VUEMNetDrives.xml in the current folder.

    .Example
    New-VUEMNetDrivesXml -SelfHealingEnabled -Disable

    Description

    -----------

    Create WEM Actions from all the Drive Mappings for the current user and export this to VUEMNetDrives.xml
    in the current folder.
    Actions are created using the Enable SelfHealing switch, and are disabled.

    .Example
    New-VUEMNetDrivesXml -OverrideEmptyDescription

    Description

    -----------

    Create WEM Actions from all the Drive Mappings for the current user and export this to VUEMNetDrives.xml
    in the current folder.
    If no Description was found during processing, the Action name is used as description.

    .Notes
    Author:  Arjan Mensch
    Version: 1.0.0

    Credentials are skipped.
#>
Function New-VUEMNetDrivesXml {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$DriveLetter = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputFileName = "VUEMNetDrives.xml",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$InputCsv = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$Prefix = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$SelfHealingEnabled = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$OverrideEmptyDescription = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$Disable = $False
    )
    Write-Verbose "$(Get-Date -Format G)"
    Write-Verbose "Function: $($MyInvocation.MyCommand)"

    # check if $DriveLetter is valid
    If ($DriveLetter -and !(Test-Path -Path "HKCU:Network\$DriveLetter")) {
        Write-Host "Cannot find '$DriveLetter' because it does not exist or is not a valid driveletter." -ForegroundColor Red
        Break
   }

   # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path -Path $OutputPath) -or ((Get-Item -Path $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    } Elseif ($OutputPath.EndsWith("\")) { $OutputPath = $OutputPath.Substring(0,$OutputPath.Length-1) }

    # check if $OutputFileName is valid
    If ($OutputFileName -and (!($OutputFileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -eq -1))) {
        Write-Host "'$OutputFileName' is not a valid filename." -ForegroundColor Red
        Break
    }

    # check if $InputFile is used and valid
    If ($InputCSV) {
        If (!(Test-Path -Path $InputCsv)) {
            Write-Host "Cannot find '$InputCsv' because it does not exist or is not a valid file." -ForegroundColor Red
            Break
        }
        If (!((Get-Content $InputCsv)[0] -split(',')).Contains("TargetPath")) {
            Write-Host "Couldn't find colum 'TargetPath' in '$InputCsv'." -ForegroundColor Red
            Break
        }
    }

    # Verbose output
    Write-Verbose "Using DriveLetter '$DriveLetter'"
    Write-Verbose "Using OutputPath '$OutputPath'"
    Write-Verbose "Using OutputFileName: '$OutputFileName'"
    Write-Verbose "Using InputCsv '$InputCsv'"
    Write-Verbose "Using Prefix '$Prefix'"
    Write-Verbose "Using SelfHealingEnabled: $SelfHealingEnabled"
    Write-Verbose "Using OverrideEmptyDescription: $OverrideEmptyDescription"
    Write-Verbose "Using Disable: $Disable"
    
    # init VUEM action arrays
    $VUEMNetDrives = @()

    # define state
    $State = "1"
    If ($Disable) { $State = "0" }

    # define selfhealing
    $SelfHeal = "0"
    If ($SelfHealingEnabled) { $SelfHeal = "1" }

    # grab mapped drives
    $MappedDrives = @()
    
    # tell user we're processing Network Drives
    Write-Host "`nProcessing Mapped Drives" -ForegroundColor Yellow

    # process using the CSV if requested
    If ($InputCsv) {
        $MappedDrivesCsv = Import-Csv -Path $InputCsv

        Write-Verbose "Found $($MappedDrivesCsv.Count) Mapped Drives to process"

        # process inputs
        ForEach ($MappedDrive in $MappedDrivesCSV) {
            $DriveName = "$Prefix$($MappedDrive.TargetPath)"
            $MappedDriveLabel = $MappedDrive.DisplayName
            If ($MappedDriveLabel) { $DriveName += " ($MappedDriveLabel)" }
                    
            $MappedDriveName = Get-UniqueActionName -ObjectList $VUEMNetDrives -ActionName $DriveName

            $Description = $null
            If ($OverrideEmptyDescription) { $Description = "$($MappedDrive.TargetPath)" }

            $VUEMNetDrive = New-VUEMNetDriveObject -Name "$MappedDriveName" `
                                                -Description "$Description" `
                                                -DisplayName "$MappedDriveLabel" `
                                                -TargetPath "$($MappedDrive.TargetPath)" `
                                                -SelfHealingEnabled "$SelfHeal" `
                                                -State "$State" `
                                                -ObjectList $VUEMNetDrives
        
            If ($VUEMNetDrive) { $VUEMNetDrives += $VUEMNetDrive }
        }
    } Else {
        If ($DriveLetter) {
            $MappedDrivesRegistry = Get-ChildItem -Path "HKCU:Network" | Where-Object { $_.PSChildName -like $DriveLetter }
        } Else {
            $MappedDrivesRegistry = Get-ChildItem -Path "HKCU:Network"
        }

        ForEach ($MappedDrive in $MappedDrivesRegistry) {
            $MappedDrives += Get-ItemProperty "HKCU:Network\$($MappedDrive.PSChildName)"
        }

        Write-Verbose "Found $($MappedDrives.Count) Mapped Drives to process"

        # process inputs
        ForEach ($MappedDrive in $MappedDrives) {
            $DriveName = "$Prefix$($MappedDrive.RemotePath)"
            $MappedDriveLabel = (Get-ItemProperty -Path "HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($MappedDrive.RemotePath.Replace('\','#'))")._LabelFromReg
            If ($MappedDriveLabel) { $DriveName += " ($MappedDriveLabel)" }
                    
            $MappedDriveName = Get-UniqueActionName -ObjectList $VUEMNetDrives -ActionName $DriveName

            $Description = $null
            If ($OverrideEmptyDescription) { $Description = "$($MappedDrive.RemotePath)" }

            $VUEMNetDrive = New-VUEMNetDriveObject -Name "$MappedDriveName" `
                                                -Description "$Description" `
                                                -DisplayName "$MappedDriveLabel" `
                                                -TargetPath "$($MappedDrive.RemotePath)" `
                                                -SelfHealingEnabled "$SelfHeal" `
                                                -State "$State" `
                                                -ObjectList $VUEMNetDrives
        
            If ($VUEMNetDrive) { $VUEMNetDrives += $VUEMNetDrive }
        }
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
    Builds an .xml file containing WEM Action definitions for Mapped Printers for the current user.
    This function creates the file containing the Actions ready for import into WEM.

    .Link
    https://msfreaks.wordpress.com

    .Parameter PrintServer
    Will only process the Mapped Printers connected to the Print Server as defined in the value of PrintServer.

    .Parameter OutputPath
    Location where the output xml file will be written to. Defaults to current folder if omitted.

    .Parameter OutputFileName
    The default filename is VUEMPrinters.xml. Use this parameter to override this if needed. 

    .Parameter InputCsv
    A csv file containing at least one field: TargetPath (Mandatory). For each TargetPath a Printer object will
    be created and written to the output xml file.
    Parameter PrintServer will be ignored if this parameter is used.

    .Parameter Prefix
    Provide a prefix string used to generate Action names (as displayed in the WEM console).

    .Parameter SelfHealingEnabled
    If used will create Actions using the SelfHealingEnabled parameter. Defaults to $False if omitted.

    .Parameter OverrideEmptyDescription
    If used will generate a description based on the Action name, but only if a description is not found
    during processing. Defaults to $False if omitted.

    .Parameter Disable
    If used will create disabled Actions. Defaults to $False if omitted (create Enabled Actions).

    .Example
    New-VUEMPrintersXml

    Description

    -----------

    Create WEM Actions from all the Printer Mappings for the current user and export this to VUEMPrinters.xml
    in the current folder.

    .Example
    New-VUEMPrintersXml -Prefix "ITW - "

    Description

    -----------

    Create WEM Actions from all the Printer Mappings for the current user and export this to VUEMPrinters.xml
    in the current folder.
    All Action names in the WEM Console are prefixed with "ITW - ", like for example "ITW - \\PrintServer\printer1".

    .Example
    New-VUEMPrintersXml -PrintServer "itwmaster"

    Description

    -----------

    Create a WEM Action for only the Printer Mappings that point to $PrintServer for the current user and export
    this to VUEMPrinters.xml in the current folder.

    .Example
    New-VUEMPrintersXml -OutputPath "C:\Temp" -OutputFileName "printers.xml"

    Description

    -----------

    Create printers.xml in c:\temp containing WEM Actions for all the Printer Mappings for the current user.

    .Example
    New-VUEMPrintersXml -InputCSV "C:\Temp\mappedprinters.csv"

    Description

    -----------

    Create WEM Actions from all the items in C:\Temp\mappedprinters.csv and export these as Printer objects to
    VUEMPrinters.xml in the current folder.

    .Example
    New-VUEMPrintersXml -SelfHealingEnabled -Disable

    Description

    -----------

    Create WEM Actions from all the Printer Mappings for the current user and export this to VUEMPrinters.xml
    in the current folder.
    Actions are created using the Enable SelfHealing switch, and are disabled.

    .Example
    New-VUEMPrintersXml -OverrideEmptyDescription

    Description

    -----------

    Create WEM Actions from all the Printer Mappings for the current user and export this to VUEMPrinters.xml
    in the current folder.
    If no Description was found during processing, the Action name is used as description.

    .Notes
    Author:  Arjan Mensch
    Version: 1.0.0

    Credentials are skipped.
#>
Function New-VUEMPrintersXml {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False,ValueFromPipeline=$False,Position=1)]
        [string]$PrintServer = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputFileName = "VUEMPrinters.xml",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$InputCsv = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$Prefix = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$SelfHealingEnabled = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$OverrideEmptyDescription = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$Disable = $False
    )
    Write-Verbose "$(Get-Date -Format G)"
    Write-Verbose "Function: $($MyInvocation.MyCommand)"

    # check if current user has mapped printers
    If (!(Test-Path -Path "HKCU:Printers\Connections")) {
        Write-Host "No Printer Mappings found in the current user context." -ForegroundColor Red
        Break
    }

    # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path -Path $OutputPath) -or ((Get-Item -Path $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    } Elseif ($OutputPath.EndsWith("\")) { $OutputPath = $OutputPath.Substring(0,$OutputPath.Length-1) }

    # check if $OutputFileName is valid
    If ($OutputFileName -and (!($OutputFileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -eq -1))) {
        Write-Host "'$OutputFileName' is not a valid filename." -ForegroundColor Red
        Break
    }

    # check if $InputFile is used and valid
    If ($InputCSV) {
        If (!(Test-Path -Path $InputCsv)) {
            Write-Host "Cannot find '$InputCsv' because it does not exist or is not a valid file." -ForegroundColor Red
            Break
        }
        If (!((Get-Content $InputCsv)[0] -split(',')).Contains("TargetPath")) {
            Write-Host "Couldn't find colum 'TargetPath' in '$InputCsv'." -ForegroundColor Red
            Break
        }
    }
    
    # Verbose output
    Write-Verbose "Using PrintServer '$PrintServer'"
    Write-Verbose "Using OutputPath '$OutputPath'"
    Write-Verbose "Using OutputFileName: '$OutputFileName'"
    Write-Verbose "Using InputCsv '$InputCsv'"
    Write-Verbose "Using Prefix '$Prefix'"
    Write-Verbose "Using SelfHealingEnabled: $SelfHealingEnabled"
    Write-Verbose "Using OverrideEmptyDescription: $OverrideEmptyDescription"
    Write-Verbose "Using Disable: $Disable"
    
    # init VUEM action arrays
    $VUEMPrinters = @()

    # define state
    $State = "1"
    If ($Disable) { $State = "0" }

    # define selfhealing
    $SelfHeal = "0"
    If ($SelfHealingEnabled) { $SelfHeal = "1" }

    # grab mapped printers
    $MappedPrinters = @()
    
    # tell user we're processing Mapped Printers
    Write-Host "`nProcessing Mapped Printers" -ForegroundColor Yellow

    # process using the CSV if requested
    If ($InputCsv) {
        $MappedPrintersCsv = Import-Csv -Path $InputCsv

        Write-Verbose "Found $($MappedPrintersCsv.Count) Mapped Printers to process"

        # process inputs
        ForEach ($MappedPrinter in $MappedPrintersCSV) {
            $PrinterName = "$Prefix$($MappedPrinter.TargetPath)"

            $MappedPrinterName = Get-UniqueActionName -ObjectList $VUEMPrinters -ActionName $PrinterName

            $Description = $null
            If ($OverrideEmptyDescription) { $Description = "$($MappedPrinter.TargetPath)" }

            $VUEMPrinter = New-VUEMPrinterObject -Name "$MappedPrinterName" `
                                                 -Description "$Description" `
                                                 -TargetPath "$($MappedPrinter.TargetPath)" `
                                                 -SelfHealingEnabled "$SelfHeal" `
                                                 -State "$State" `
                                                 -ObjectList $VUEMPrinters
        
            If ($VUEMPrinter) { $VUEMPrinters += $VUEMPrinter }
        }
    } Else {
        If ($PrintServer) {
            $MappedPrintersRegistry = Get-ChildItem -Path "HKCU:Printers\Connections" | Where-Object { $_.PSChildName -like "*$($PrintServer)*" }
        } Else {
            $MappedPrintersRegistry = Get-ChildItem -Path "HKCU:Printers\Connections"
        }
        ForEach ($MappedPrinter in $MappedPrintersRegistry) {
            $MappedPrinters += Get-ItemProperty -Path "HKCU:Printers\Connections\$($MappedPrinter.PSChildName)"
        }

        Write-Verbose "Found $($MappedPrinters.Count) Mapped Printers to process"

        # process inputs
        ForEach ($MappedPrinter in $MappedPrinters) {
            $PrinterName = "$Prefix$($MappedPrinter.PSChildName.Replace(",","\"))"

            $MappedPrinterName = Get-UniqueActionName -ObjectList $VUEMPrinters -ActionName $PrinterName

            $Description = $null
            If ($OverrideEmptyDescription) { $Description = "$($MappedPrinter.PSChildName.Replace(",","\"))" }

            $VUEMPrinter = New-VUEMPrinterObject -Name "$MappedPrinterName" `
                                                -Description "$Description" `
                                                -TargetPath "$($MappedPrinter.PSChildName.Replace(",","\"))" `
                                                -SelfHealingEnabled "$SelfHeal" `
                                                -State "$State" `
                                                -ObjectList $VUEMPrinters
        
            If ($VUEMPrinter) { $VUEMPrinters += $VUEMPrinter }
        }
    }

    # output xml file
    If ($VUEMPrinters) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMPrinter" -ObjectList $VUEMPrinters | Out-File $OutputPath\$OutputFileName
        Write-Host "$OutputFileName written to '$OutputPath\$OutputFileName'" -ForegroundColor Green
    }
}

<# 
    .Synopsis
    Builds an .xml file containing WEM Action definitions.

    .Description
    Builds an .xml file containing WEM Action definitions for UserDSN entries.
    This function creates the xml file containing the Actions ready for import into WEM.

    .Link
    https://msfreaks.wordpress.com

    .Parameter Name
    Will only process the DSN defined by $Name.

    .Parameter OutputPath
    Location where the output xml file will be written to. Defaults to current folder if omitted.

    .Parameter OutputFileName
    The default filename is VUEMUserDSNs.xml. Use this parameter to override this if needed. 

    .Parameter SystemDSN
    If this parameter is used, the script will process SystemDSN into UserDSN Actions.

    .Parameter Prefix
    Provide a prefix string used to generate Action names (as displayed in the WEM console).

    .Parameter RunOnce
    If used will create Actions using the RunOnce parameter. Defaults to $False if omitted.

    .Parameter OverrideEmptyDescription
    If used will generate a description based on the Action name, but only if a description is not found
    during processing. Defaults to $False if omitted.

    .Parameter Disable
    If used will create disabled Actions. Defaults to $False if omitted (create Enabled Actions).

    .Example
    New-VUEMUserDSNsXml

    Description

    -----------

    Create WEM Actions from all the User DSNs for the current user and export this to VUEMUserDSNs.xml in the
    current folder.

    .Example
    New-VUEMUserDSNsXml -Name "WEM Database Connection" -System -RunOnce

    Description

    -----------

    Create WEM UserDSN Action from the "WEM Database Connection" for either the current user, or for the
    local system.
    Eventhough this might be a System DSN, the script will create a UserDSN action.
    The RunOnce parameter will be enabled for these actions because the RunOnce switch is used.

    .Example
    New-VUEMUserDSNsXml -Prefix "ITW - "

    Description

    -----------

    Create WEM Actions from all the User DSNs for the current user and export this to VUEMUserDSNs.xml in the
    current folder.
    All Action names in the WEM Console are prefixed with "ITW - ", like for example "ITW - WEM Database".

    .Example
    New-VUEMUserDSNsXml -OutputPath "C:\Temp" -OutputFileName "dsns.xml" -Disable

    Description

    -----------

    Create dsns.xml in c:\temp for all the User DSNs for the current user. The Actions will be disabled once
    imported into WEM.

    .Example
    New-VUEMUserDSNsXml -OverrideEmptyDescription

    Description

    -----------

    Create WEM Actions from all the User DSNs for the current user and export this to VUEMUserDSNs.xml in the
    current folder.
    If no Description was found during processing, the Action name is used as description.

    .Notes
    Author:  Arjan Mensch
    Version: 1.0.0

    Seems WEM only supports DSNs based on the "SQL Server" driver, so all DataSources based on other drivers
    are skipped.
    Credentials are skipped.
#>
Function New-VUEMUserDSNsXml {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False,ValueFromPipeline=$False,Position=1)]
        [string]$Name = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputPath = (Resolve-Path .\).Path,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$OutputFileName = "VUEMUserDSNs.xml",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$SystemDSN = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [string]$Prefix = "",
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$RunOnce = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$OverrideEmptyDescription = $False,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
        [switch]$Disable = $False
    )
    Write-Verbose "$(Get-Date -Format G)"
    Write-Verbose "Function: $($MyInvocation.MyCommand)"

    # check if $Name is valid
    If (($Name -and $SystemDSN -and (!(Get-OdbcDsn | Where-Object { $_.Name -like "$Name" }))) `
        -or ($Name -and !($SystemDSN) -and (!(Get-OdbcDsn | Where-Object { $_.Name -like "$Name" -and $_.DsnType -eq "User" })))) {
        Write-Host "Cannot find User DSN or System DSN '$Name' because it does not exist." -ForegroundColor Red
        Break
    }

    # check if $OutputPath is valid
    If ($OutputPath -and (!(Test-Path -Path $OutputPath) -or ((Get-Item -Path $OutputPath) -isnot [System.IO.DirectoryInfo]))) {
        Write-Host "Cannot find path '$OutputPath' because it does not exist or is not a valid path." -ForegroundColor Red
        Break
    } Elseif ($OutputPath.EndsWith("\")) { $OutputPath = $OutputPath.Substring(0,$OutputPath.Length-1) }

    # check if $OutputFileName is valid
    If ($OutputFileName -and (!($OutputFileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -eq -1))) {
        Write-Host "'$OutputFileName' is not a valid filename." -ForegroundColor Red
        Break
    }

    # Verbose output
    Write-Verbose "Using Name '$Name'"
    Write-Verbose "Using OutputPath '$OutputPath'"
    Write-Verbose "Using OutputFileName: '$OutputFileName'"
    Write-Verbose "Using SystemDSN $SystemDSN"
    Write-Verbose "Using Prefix '$Prefix'"
    Write-Verbose "Using RunOnce: $RunOnce"
    Write-Verbose "Using OverrideEmptyDescription: $OverrideEmptyDescription"
    Write-Verbose "Using Disable: $Disable"
    
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
    Write-Verbose "Found $(If ($dsns.Count) { $dsns.Count } else { "1" } ) DSNs to process"

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

        $Description = "$($dsn.Attribute.Description)".Trim()
        If (!$Description -and $OverrideEmptyDescription) { $Description = "$($dsn.Name)" }

        $VUEMUserDSN = New-VUEMUserDSNObject -Name "$VUEMDSNName" `
                                             -Description "$Description" `
                                             -TargetName "$($dsn.Name)" `
                                             -TargetDriverName "$($dsn.DriverName)" `
                                             -TargetServerName "$($dsn.Attribute.Server)" `
                                             -TargetDatabaseName "$($dsn.Attribute.Database)" `
                                             -RunOnce "$VUEMRunOnce" `
                                             -State "$State" `
                                             -ObjectList $VUEMUserDSNs
    
        If ($VUEMUserDSN) { $VUEMUserDSNs += $VUEMUserDSN }
    }

    # output xml file
    If ($VUEMUserDSNs) {
        New-VUEMXmlFile -VUEMIdentifier "VUEMUserDSN" -ObjectList $VUEMUserDSNs | Out-File $OutputPath\$OutputFileName
        Write-Host "$OutputFileName written to '$OutputPath\$OutputFileName'" -ForegroundColor Green
    }
}

#region Helper Functions (will not be exposed when module is loaded)
<#
 .SYNOPSIS
  Helper function to grab an array of files to process.
#>
Function Get-FilesToProcess{
    param(
        [Parameter(Mandatory=$True)][string]$Path
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
        If ($Recurse) {
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
Function Get-IconToBase64{
    param(
        [Parameter(Mandatory=$True)][object]$Icon
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
Function New-VUEMApplicationObject() {
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
        [string]$State,
        [psobject[]]$ObjectList
    )

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Description -eq $Description `
                                                       -and $_.DisplayName -eq $DisplayName `
                                                       -and $_.StartMenuTarget -eq $StartMenuTarget `
                                                       -and $_.TargetPath -eq $TargetPath `
                                                       -and $_.Parameters -eq $Parameters `
                                                       -and $_.WorkingDirectory -eq $WorkingDirectory `
                                                       -and $_.Hotkey -eq $Hotkey `
                                                       -and $_.IconLocation -eq $IconLocation `
                                                       -and $_.IconStream -eq $IconStream `
                                                       -and $_.Reserved01 -eq '<?xml version="1.0" encoding="utf-8"?><ArrayOfVUEMActionAdvancedOption xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><VUEMActionAdvancedOption><Name>SelfHealingEnabled</Name><Value>'+$SelfHealingEnabled+'</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>EnforceIconLocation</Name><Value>0</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>EnforcedIconXValue</Name><Value>0</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>EnforcedIconYValue</Name><Value>0</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>DoNotShowInSelfService</Name><Value>0</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>CreateShortcutInUserFavoritesFolder</Name><Value>0</Value></VUEMActionAdvancedOption></ArrayOfVUEMActionAdvancedOption>' `
                                                       -and $_.State -eq $State })) {

        Write-Verbose " Skipped Application action '$Name', already in array"
        Return $null
    }

    Write-Verbose " Added Application action '$Name'"

    $AppType = "0"
    If ($WorkingDirectory -like "Url") { $AppType = "4" }
    Return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'DisplayName' = $DisplayName
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
        'AppType' = $AppType
        'State' = $State
    }
    # Action type 0 = create application shortcut
}

<#
 .SYNOPSIS
  Helper function to create VUEMExtTask object
#>
Function New-VUEMExtTaskObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$TargetPath,
        [string]$TargetArgs,
        [string]$State,
        [psobject[]]$ObjectList
    )

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Description -eq $Description `
                                                      -and $_.TargetPath -eq $TargetPath `
                                                      -and $_.TargetArgs -eq $TargetArgs `
                                                      -and $_.State -eq $State })) {

        Write-Verbose " Skipped External Task '$Name', already in array"
        Return $null
    }

    Write-Verbose " Added External Task '$Name'"

    Return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'TargetPath' = $TargetPath
        'TargetArgs' = $TargetArgs
        'Reserved01' = '<?xml version="1.0" encoding="utf-8"?><ArrayOfVUEMActionAdvancedOption xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><VUEMActionAdvancedOption><Name>ExecuteOnlyAtLogon</Name><Value>1</Value></VUEMActionAdvancedOption></ArrayOfVUEMActionAdvancedOption>'
        'TimeOut' = "30"
        'WaitForFinish' = "0"
        'ExecOrder' = "0"
        'RunHidden' = "1"
        'RunOnce' = "0"
        'ActionType' = "0"
        'State' = $State
    }
    # Action type 0 = create external task
}

<#
 .SYNOPSIS
  Helper function to create VUEMNetDrive object
#>
Function New-VUEMNetDriveObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$DisplayName,
        [string]$TargetPath,
        [string]$SelfHealingEnabled,
        [string]$State,
        [psobject[]]$ObjectList
    )

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Description -eq $Description `
                                                       -and $_.DisplayName -eq $DisplayName `
                                                       -and $_.TargetPath -eq $TargetPath `
                                                       -and $_.Reserved01 -eq '<?xml version="1.0" encoding="utf-8"?><ArrayOfVUEMActionAdvancedOption xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><VUEMActionAdvancedOption><Name>SelfHealingEnabled</Name><Value>'+$SelfHealingEnabled+'</Value></VUEMActionAdvancedOption><VUEMActionAdvancedOption><Name>SetAsHomeDriveEnabled</Name><Value>0</Value></VUEMActionAdvancedOption></ArrayOfVUEMActionAdvancedOption>' `
                                                       -and $_.State -eq $State })) {

        Write-Verbose " Skipped Drive Mapping '$Name', already in array"
        Return $null
    }

    Write-Verbose " Added Drive Mapping '$Name'"

    Return [pscustomobject] @{     
        'Name' = $Name
        'Description' = $Description
        'DisplayName' = $DisplayName
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
Function New-VUEMEnvVariableObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$VariableName,
        [string]$VariableValue,
        [string]$VariableType,
        [string]$State,
        [psobject[]]$ObjectList
    )

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Description -eq $Description `
                                                       -and $_.VariableName -eq $VariableName `
                                                       -and $_.VariableValue -eq $VariableValue `
                                                       -and $_.VariableType -eq $VariableType `
                                                       -and $_.State -eq $State })) {

        Write-Verbose " Skipped Environment Variable '$Name', already in array"
        Return $null
    }

    Write-Verbose " Added Environment Variable '$Name'"

    Return [pscustomobject] @{     
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
Function New-VUEMFileSystemOpObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$SourcePath,
        [string]$TargetPath,
        [string]$RunOnce,
        [string]$ActionType,
        [string]$State,
        [psobject[]]$ObjectList
    )

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Description -eq $Description `
                                                       -and $_.SourcePath -eq $SourcePath `
                                                       -and $_.TargetPath -eq $TargetPath `
                                                       -and $_.RunOnce -eq $RunOnce -and $_.ActionType -eq $ActionType `
                                                       -and $_.State -eq $State })) {

        Write-Verbose " Skipped FileSystem action '$Name', already in array"
        Return $null
    }

    Write-Verbose " Added FileSystem action '$Name'"

    Return [pscustomobject] @{     
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
Function New-VUEMIniFileOpObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$TargetPath,
        [string]$TargetSectionName,
        [string]$TargetValueName,
        [string]$TargetValue,
        [string]$RunOnce,
        [string]$State,
        [psobject[]]$ObjectList
    )

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Description -eq $Description `
                                                       -and $_.TargetPath -eq $TargetPath `
                                                       -and $_.TargetSectionName -eq $TargetSectionName `
                                                       -and $_.TargetValueName -eq $TargetValueName `
                                                       -and $_.TargetValue -eq $TargetValue `
                                                       -and $_.RunOnce -eq $RunOnce -and $_.State -eq $State })) {

        Write-Verbose " Skipped IniFile action '$Name', already in array"
        Return $null
    }

    Write-Verbose " Added IniFile action '$Name'"

    Return [pscustomobject] @{     
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
Function New-VUEMPrinterObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$TargetPath,
        [string]$SelfHealingEnabled,
        [string]$State,
        [psobject[]]$ObjectList
    )

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Description -eq $Description `
                                                       -and $_.TargetName -eq $TargetName `
                                                       -and $_.TargetPath -eq $TargetPath `
                                                       -and $_.Reserved01 -eq '<?xml version="1.0" encoding="utf-8"?><ArrayOfVUEMActionAdvancedOption xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><VUEMActionAdvancedOption><Name>SelfHealingEnabled</Name><Value>'+$SelfHealingEnabled+'</Value></VUEMActionAdvancedOption></ArrayOfVUEMActionAdvancedOption>' `
                                                       -and $_.State -eq $State })) {

        Write-Verbose " Skipped Printer Mapping '$Name', already in array"
        Return $null
    }

    Write-Verbose " Added Printer Mapping '$Name'"

    Return [pscustomobject] @{     
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
    # Action type 0 = map printer 
}

<#
 .SYNOPSIS
  Helper function to create VUEMRegValue object
#>
Function New-VUEMRegValueObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$TargetName,
        [string]$TargetPath,
        [string]$TargetType,
        [string]$TargetValue,
        [string]$RunOnce,
        [String]$ActionType,
        [string]$State,
        [psobject[]]$ObjectList
    )
    
    If (!$TargetName) { 
        $TargetName = "(Default)"
        $TargetType = "REG_SZ"
    }

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Description -eq $Description `
                                                       -and $_.TargetName -eq $TargetName `
                                                       -and $_.TargetPath -eq $TargetPath `
                                                       -and $_.TargetType -eq $TargetType `
                                                       -and $_.TargetValue -eq $TargetValue `
                                                       -and $_.RunOnce -eq $RunOnce -and $_.ActionType -eq $ActionType `
                                                       -and $_.State -eq $State })) {

        Write-Verbose " Skipped Registry entry '$Name', already in array"
        Return $null
    }

    Write-Verbose " Added Registry entry '$Name'"

    Return [pscustomobject] @{     
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
Function New-VUEMUserDSNObject() {
    param(
        [string]$Name,
        [string]$Description,
        [string]$TargetName,
        [string]$TargetDriverName,
        [string]$TargetServerName,
        [string]$TargetDatabaseName,
        [string]$RunOnce,
        [string]$State,
        [psobject[]]$ObjectList
    )

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Description -eq $Description `
                                                       -and $_.TargetName -eq $TargetName `
                                                       -and $_.TargetDriverName -eq $TargetDriverName `
                                                       -and $_.TargetServerName -eq $TargetServerName `
                                                       -and $_.TargetDatabaseName -eq $TargetDatabaseName `
                                                       -and $_.RunOnce -eq $RunOnce -and $_.State -eq $State })) {

        Write-Verbose " Skipped UserDSN '$Name', already in array"
        Return $null
    }

    Write-Verbose " Added UserDSN '$Name'"

    Return [pscustomobject] @{     
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
  Helper function to create GPOFilter object
#>
Function New-GPOFilterObject() {
    param(
        [string]$Name,
        [string]$ActionType,
        [object]$Filter,
        [psobject[]]$ObjectList
    )

    # if $Filter is $null return nothing
    If (!$Filter) { Return $null }

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Name -eq $Name `
                                                       -and $_.ActionType -eq $ActionType `
                                                       -and $_.Filter -eq $Filter })) {

        Write-Verbose " Skipping '$Name' filter, already in array"
        Return $null
    }

    Write-Verbose " Found filter '$($Filter.InnerXml)'"

    Return [pscustomobject] @{     
        'Name' = $Name
        'ActionType' = $ActionType
        'Filter' = $Filter.InnerXml
    }
}

<#
 .SYNOPSIS
  Helper function for creating a VUEM xml file
#>
Function New-VUEMXmlFile {
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
            If ($p.Name -ne "Init") { $xmlWriter.WriteElementString($p.Name,$p.Value) }
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
  Helper function to create VUEMConfigurationSetting object
#>
Function New-VUEMConfigurationSettingObject() {
    param(
        [string]$State,
        [string]$Type,
        [string]$Name,
        [string]$Value,
        [bool]$Init = $False,
        [psobject[]]$ObjectList
    )

    # check if object is unique
    If ($ObjectList -and ($ObjectList | Where-Object { $_.Name -eq $Name -and !$_.Init})) {
        Write-Verbose " Skipping '$Name' setting, newer instance found"
        Return $ObjectList
    }

    # if object was in init state, remove it
    If ($ObjectList -and !$Init -and ($ObjectList | Where-Object { $_.Name -eq $Name -and $_.Init})) {
        $ObjectList = $ObjectList | Where-Object { $_.Name -notlike $Name }
    }

    $VUEMConfigurationSetting = `
    [pscustomobject] @{     
        'State' = $State
        'Type' = $Type
        'Name' = $Name
        'Value' = $Value
        'Reserved01' = $null
        'Init' = $Init
    }

    $ObjectList += $VUEMConfigurationSetting
    Write-Verbose " Added '$Name' setting"

    Return $ObjectList
}

<#
 .SYNOPSIS
  Helper function to create VUEMVUSVConfigurationSettings initial objectlist
#>
Function Get-VUEMMicrosoftUsvInitialObjects {
    param(
        [string]$Value = "0"
    )

    $ObjectList = @()
    # default mandatory objects
    $ObjectList += [pscustomobject] @{ "Name"="processUSVConfiguration"; "State"="1"; "Type"="0"; "Value"="$Value"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processUSVConfigurationForAdmins"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    # tab Roaming Profiles Configuration
    $ObjectList += [pscustomobject] @{ "Name"="SetWindowsRoamingProfilesPath"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="WindowsRoamingProfilesPath"; "State"="1"; "Type"="1"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SetRDSRoamingProfilesPath"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="RDSRoamingProfilesPath"; "State"="1"; "Type"="1"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SetRDSHomeDrivePath"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="RDSHomeDrivePath"; "State"="1"; "Type"="1"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="RDSHomeDriveLetter"; "State"="1"; "Type"="1"; "Value"="Z:"; "Reserved01"=""; "Init"=$True }
    #tab Roaming Profiles Advanced Configuration
    $ObjectList += [pscustomobject] @{ "Name"="SetRoamingProfilesFoldersExclusions"; "State"="1"; "Type"="2"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="RoamingProfilesFoldersExclusions"; "State"="1"; "Type"="2"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DeleteRoamingCachedProfiles"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="AddAdminGroupToRUP"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="CompatibleRUPSecurity"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisableSlowLinkDetect"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SlowLinkProfileDefault"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    # tab Folder Redirection
    $ObjectList += [pscustomobject] @{ "Name"="processFoldersRedirectionConfiguration"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DeleteLocalRedirectedFolders"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processDesktopRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DesktopRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processStartMenuRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="StartMenuRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processPersonalRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="PersonalRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processPicturesRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="PicturesRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="MyPicturesFollowsDocuments"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processMusicRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="MusicRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="MyMusicFollowsDocuments"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processVideoRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="VideoRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="MyVideoFollowsDocuments"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processFavoritesRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="FavoritesRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processAppDataRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="AppDataRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processContactsRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="ContactsRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processDownloadsRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DownloadsRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processLinksRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="LinksRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processSearchesRedirection"; "State"="1"; "Type"="3"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SearchesRedirectedPath"; "State"="1"; "Type"="3"; "Value"=""; "Reserved01"=""; "Init"=$True }

    Return $ObjectList
}

<#
 .SYNOPSIS
  Helper function to create VUEMEnvironmentalSettings initial objectlist
#>
Function Get-VUEMEnvironmentalInitialObjects {
    param(
        [string]$Value = "0"
    )

    $ObjectList = @()
    # default mandatory objects
    $ObjectList += [pscustomobject] @{ "Name"="processEnvironmentalSettings"; "State"="1"; "Type"="2"; "Value"="$Value"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="processEnvironmentalSettingsForAdmins"; "State"="1"; "Type"="2"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    # tab Start Menu
    $ObjectList += [pscustomobject] @{ "Name"="HideCommonPrograms"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="RemoveRunFromStartMenu"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideAdministrativeTools"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideHelp"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideFind"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideWindowsUpdate"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="LockTaskbar"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideSystemClock"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideDevicesandPrinters"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideTurnOff"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="ForceLogoff"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="Turnoffnotificationareacleanup"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="TurnOffpersonalizedmenus"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="ClearRecentprogramslist"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SetSpecificThemeFile"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SpecificThemeFileValue"; "State"="1"; "Type"="1"; "Value"="%windir%\resources\Themes\aero.theme"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SetVisualStyleFile"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="VisualStyleFileValue"; "State"="1"; "Type"="1"; "Value"="%windir%\resources\Themes\Aero\aero.msstyles"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SetWallpaper"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="Wallpaper"; "State"="1"; "Type"="1"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="WallpaperStyle"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SetDesktopBackGroundColor"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DesktopBackGroundColor"; "State"="1"; "Type"="0"; "Value"=""; "Reserved01"=""; "Init"=$True }
    # tab Desktop
    $ObjectList += [pscustomobject] @{ "Name"="NoMyComputerIcon"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="NoRecycleBinIcon"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="NoMyDocumentsIcon"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="BootToDesktopInsteadOfStart"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="NoPropertiesMyComputer"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="NoPropertiesRecycleBin"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="NoPropertiesMyDocuments"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideNetworkIcon"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideNetworkConnections"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisableTaskMgr"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisableTLcorner"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisableCharmsHint"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    # tab Windows Explorer
    $ObjectList += [pscustomobject] @{ "Name"="DisableRegistryEditing"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisableSilentRegedit"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisableCmd"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisableCmdScripts"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="RemoveContextMenuManageItem"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="NoNetConnectDisconnect"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideLibrairiesInExplorer"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideNetworkInExplorer"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="NoProgramsCPL"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="NoNtSecurity"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="NoViewContextMenu"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="NoTrayContextMenu"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="HideSpecifiedDrivesFromExplorer"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="ExplorerHiddenDrives"; "State"="1"; "Type"="1"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="RestrictSpecifiedDrivesFromExplorer"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="ExplorerRestrictedDrives"; "State"="1"; "Type"="1"; "Value"=""; "Reserved01"=""; "Init"=$True }
    # tab Control Panel
    $ObjectList += [pscustomobject] @{ "Name"="HideControlPanel"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="RestrictCpl"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="RestrictCplList"; "State"="1"; "Type"="0"; "Value"=""; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisallowCpl"; "State"="1"; "Type"="0"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisallowCplList"; "State"="1"; "Type"="0"; "Value"=""; "Reserved01"=""; "Init"=$True }
    # tab Known Folder Management
    $ObjectList += [pscustomobject] @{ "Name"="DisableSpecifiedKnownFolders"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisabledKnownFolders"; "State"="1"; "Type"="1"; "Value"=""; "Reserved01"=""; "Init"=$True }
    # tab SBC / HDV Tuning
    $ObjectList += [pscustomobject] @{ "Name"="DisableDragFullWindows"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisableCursorBlink"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="EnableAutoEndTasks"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="WaitToKillAppTimeout"; "State"="1"; "Type"="1"; "Value"="20000"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisableSmoothScroll"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="DisableMinAnimate"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SetCursorBlinkRate"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="CursorBlinkRateValue"; "State"="1"; "Type"="1"; "Value"="-1"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SetMenuShowDelay"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="MenuShowDelayValue"; "State"="1"; "Type"="1"; "Value"="10"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="SetInteractiveDelay"; "State"="1"; "Type"="1"; "Value"="0"; "Reserved01"=""; "Init"=$True }
    $ObjectList += [pscustomobject] @{ "Name"="InteractiveDelayValue"; "State"="1"; "Type"="1"; "Value"="40"; "Reserved01"=""; "Init"=$True }

    Return $ObjectList
}

<#
 .SYNOPSIS
  Helper function to ensure a unique action name
#>
Function Get-UniqueActionName {
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

<#
 .SYNOPSIS
  Helper function to recursively check collections for registrysettings
#>
Function Get-GPORegistrySettingsFromCollection {
    param(
        [object[]]$Collections,
        [System.Xml.XmlNode[]]$Filters = @()
    )

    $RegistrySettings = @()

    ForEach ($Collection in $Collections) {
        If ($Collection.Filters) { $Filters += $Collection.Filters }

        If ($Filters -and $Collection.Registry) {            
            ForEach ($Filter in $Filters) { 
                ForEach ($Registry in $Collection.Registry) {
                    $Registry.AppendChild($Filter)
                    $RegistrySettings += $Registry
                }
            }
        } Else {
            $RegistrySettings += $Collection.Registry
        }

        If ($Collection.Collection) {
            $RegistrySettings += Get-GPORegistrySettingsFromCollection -Collections ($Collection.Collection) -Filters $Filters
        }

    }

    Return $RegistrySettings
}
#endregion

# expose functions
Export-ModuleMember -Function Import-VUEMActionsFromGpo
Export-ModuleMember -Function Import-VUEMActionsFromBrokerApplicationCSV
Export-ModuleMember -Function Import-VUEMEnvironmentalSettingsFromGpo
Export-ModuleMember -Function Import-VUEMMicrosoftUsvSettingsFromGpo
Export-ModuleMember -Function New-VUEMApplicationsXml
Export-ModuleMember -Function New-VUEMNetDrivesXml
Export-ModuleMember -Function New-VUEMPrintersXml
Export-ModuleMember -Function New-VUEMUserDSNsXml
