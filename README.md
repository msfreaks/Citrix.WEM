
# Citrix.WEM
Powershell module for Citrix WEM

Detailed information for this module can be found on my Blog:
https://msfreaks.wordpress.com

I started this project after onboarding several environments using Citrix WEM.
Got real tired real fast of adding applications to the WEM Actions manually.
This module contains functions to make an admin's life much easier when onboarding
new environments in WEM.

This module quickly grew into something much bigger after all the attention it got.
It now contains seven functions in total, covering all Group Policy preferences, Environmental Settings,
Microsoft User State Virtualization Settings, and functions to run in a User Environment to mimic settings
found there into WEM.

Using the functions in this module will create .xml files you can then import (restore) into WEM.
These functions will not modify WEM or its database directly!

A big thank you to James Kindon (@james_kindon) for support, testing, suggestions, brainpicking, etc!

v1.1.0 - New functionality - January 15 2019
Added Import-VUEMActionsFromCSV: added function

v1.0.0 - Major version Release - January 8 2018
Updated all functions: -Verbose switch shows Verbose Output
Fixed Import-VUEM..FromGpo functions: Feedback on errors for GPOBackupPath parameter returned empty value
Fixed Import-VUEMActionsFromGpo: Disable parameter was not working
Fixed Import-VUEMEnvironmentalSettingsFromGpo: Enable parameter was not working
Fixed Import-VUEMEnvironmentalSettingsFromGpo: no more xml output if no settings were found
Fixed Import-VUEMMicrosoftUsvSettingsFromGpo: Enable parameter was not working
Fixed Import-VUEMMicrosoftUsvSettingsFromGpo: no more xml output if no settings were found
Fixed New-VUEM..Xml functions: OutputFileName is now checked for valid filename syntax
Fixed New-VUEM..Xml functions: Output message no longer contains '\\' if OutputPath ended in '\'

v0.9.6 - Update release - January 5 2018 (non-public internal release, aka the James Kindon release)
Updated Import-VUEMActionsFromGpo: RunOnLogonPrograms now also captured from Machine GPO as External Tasks
Updated New-VUEMNetDrivesXml: new parameter (InputCsv) to accept a csv file for input
Updated New-VUEMNetPrintersXml: new parameter (InputCsv) to accept a csv file for input
Updated all functions: new parameter (OverrideEmptyDescription)
Added Import-VUEMEnvironmentalSettingsFromGpo: added function
Added Import-VUEMMicrosoftUsvSettingsFromGpo: added function
Fixed all functions: $null object no longer added to the list of objects in Duplicate Object check
Fixed Import-VUEMActionsFromGpo: sometimes a '\' registry entry was processed in Registry Actions
Fixed Import-VUEMActionsFromGpo: Duplicate items are now dropped in Export Filters
Fixed New-VUEMApplicationObject: DisplayName variable had a typo (thank you Marcus Niemann)
Fixed New-VUEMNetDriveObject: DisplayName variable had a typo (thank you Marcus Niemann)

v0.9.5 - Update release - January 1 2018 (non-public internal release)
Updated Import-VUEMActionsFromGpo: new parameters (SelfHealingEnabled and ExportFilters)
Updated Import-VUEMActionsFromGpo: Printer processing will now detect deployed printers
Updated Import-VUEMActionsFromGpo: UserDSN processing will now process SystemDSNs
Updated Import-VUEMActionsFromGpo: GPO Filters are now captured during processing
Updated Import-VUEMActionsFromGpo: RunOnLogonPrograms now captured from GPO as External Tasks
Updated Import-VUEMActionsFromGpo: Logon Scripts now captured from GPO as External Tasks
Updated all functions: Duplicate Actions are now dropped completely
Updated NetDrive Action: name convention (now based on UNC Path)
Updated Printer Action: name convention (now based on UNC Path)
Updated RegValue Action: name convention (now based on Registry Path)
Added New-VUEMPrintersXml function
Fixed Import-VUEMActionsFromGpo: Registry Actions are now recursively processed when Collections are found

v0.9.4 - Update release - December 27 2017
Updated New-VUEMApplicationsXml: new parameters
Added New-VUEMNetDrivesXml: new function
Added New-VUEMUserDSNsXml: new function
Added Import-VUEMActionsFromGpo: new function
Fixed all functions: State was always enabled, Disable parameter was ignored

v0.9.1 - Update release - December 18 2017
Updated New-VUEMApplicationsXml

v0.9.0 - Initial release - December 17 2017
Added New-VUEMApplicationsXml: added function

Function Import-VUEMActionsFromCSV
Create a CSV using Get-BrokerApplication | Export-CSV -Path <path to output csv file>
Then use this function to import the published applications or content to WEM Application Actions.
If the target resources are available on the machine where you run this, the icons will be grabbed as well.
If you have .ico files that have the same name as the CommandLine targets for your published applications
and they are available in the same folder as your CSV file, those icons will be used.
See "Get-Help Import-VUEMActionsFromCSV" for details on how to use the function.


Function Import-VUEMActionsFromGpo
Imports User Preference settings from GPOs and converts them to WEM Action files.
This function only works on GPO Backup files, it will not communicate directly with
Active Directory to retreive the settings.
See "Get-Help Import-VUEMActionsFromGpo" for details on how to use the function.


Function Import-VUEMEnvironmentalSettingsFromGpo
Imports Environmental Settings from GPOs and converts them to WEM Environmental Settings.
This function only works on GPO Backup files, it will not communicate directly with
Active Directory to retreive the settings.
Importing the xml into WEM will override any settings already there!
See "Get-Help Import-VUEMEnvironmentalSettingsFromGpo" for details on how to use the function.


Function Import-VUEMMicrosoftUsvSettingsFromGpo
Imports Microsoft Userstate Virtualization Settings from GPOs and converts them to WEM Microsoft USV Settings.
This function only works on GPO Backup files, it will not communicate directly with
Active Directory to retreive the settings.
Importing the xml into WEM will override any settings already there!
See "Get-Help Import-VUEMMicrosoftUsvSettingsFromGpo" for details on how to use the function.


Function New-VUEMApplicationsXml
Builds an .xml file containing WEM Action definitions for application shortcuts.
This function supports multiple types of input and creates the file containing the Actions
ready for import into WEM.
See "Get-Help New-VUEMApplicationsXML" for details on how to use the function.


Function New-VUEMNetDrivesXml
Builds an .xml file containing WEM Action definitions for Mapped Network Drives for the current user.
This function creates the file containing the Actions ready for import into WEM.
See "Get-Help New-VUEMNetDrivesXML" for details on how to use the function.


Function New-VUEMPrintersXml
Builds an .xml file containing WEM Action definitions for Mapped Printers for the current user.
This function creates the file containing the Actions ready for import into WEM.
See "Get-Help New-VUEMPrintersXML" for details on how to use the function.


Function New-VUEMUserDSNsXml
Builds an .xml file containing WEM Action definitions for UserDSN entries.
This function supports multiple types of input and creates the file containing the Actions
ready for import into WEM.
See "Get-Help New-VUEMUserDSNsXML" for details on how to use the function.
