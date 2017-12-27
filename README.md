
# Citrix.WEM
Powershell module for Citrix WEM

I started this project after onboarding several environments using Citrix WEM.
Got real tired of adding applications to the WEM Actions manually.

Using the functions in this module will create .xml files you can then import (restore) into WEM.
This function will not modify WEM or its database directly!

v0.9.4 - Update release - December 27 2017
-Updated New-VUEMApplicationsXML with new parameters
-Fixed bug: State was always enabled, Disable parameter was ignored
-Added New-VUEMNetDrivesXML function
-Added New-VUEMUserDSNsXML function
-Added Import-VUEMActionsFromGPO function

v0.9.1 - Update release - December 18 2017
-Updated New-VUEMApplicationsXML

v0.9.0 - Initial release - December 17 2017
-Added New-VUEMApplicationsXML function

Function New-VUEMApplicationsXML
Builds an .xml file containing WEM Action definitions for application shortcuts.
This function supports multiple types of input and creates the file containing the Actions
ready for import into WEM.
See "Get-Help New-VUEMApplicationsXML" for details on how to use the function.


Function New-VUEMNetDrivesXML
Builds an .xml file containing WEM Action definitions for Mapped Network Drives for the current user.
This function creates the file containing the Actions ready for import into WEM.
See "Get-Help New-VUEMNetDrivesXML" for details on how to use the function.


Function New-VUEMUserDSNsXML
Builds an .xml file containing WEM Action definitions for UserDSN entries.
This function supports multiple types of input and creates the file containing the Actions
ready for import into WEM.
See "Get-Help New-VUEMUserDSNsXML" for details on how to use the function.


Function Import-VUEMActionsFromGPO
Imports User Preference settings from GPOs and converts them to WEM Action files.
This function only works on GPO Backup files, it will not communicate directly with
Active Directory to retreive the settings.
See "Get-Help Import-VUEMActionsFromGPO" for details on how to use the function.
