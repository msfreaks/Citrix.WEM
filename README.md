# Citrix.WEM
Powershell module for Citrix WEM

Started this project after onboarding several environments using Citrix WEM.
Got real tired of adding applications to the WEM Actions manually.

Right now the module contains one function, and that function does just that:
Add application shortcuts to the Actions as defined in WEM.
Function can process a single file, or a file / folder structure.

Using this function will create an .xml file you can then import (restore) into WEM.
This function will not modify WEM or its database directly!
