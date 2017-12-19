<# 
 .Synopsis
  Builds an .xml file containing WEM Action definitions.

 .Description
  Builds an .xml file containing WEM Action definitions.
  This function supports multiple types of input and creates the file containing the Actions
  ready for import into WEM.

 .Link
  https://msfreaks.wordpress.com

 .Parameter Path
  Can be a targetted folder or a targetted file. If a folder is specified the function
  will assume you wish to parse .lnk files.
  If ommitted the default Start Menu\Programs folders (current user and all users)
  locations will be used.

 .Parameter Recurse
  Whether the script needs to recursively process $Path. Only valid when the $Path parameter
  is a folder. Is $True by default if ommitted.

 .Parameter FileTypes
  Provide a comma seperated list of filetypes to process. Only valid when the $Path
  parameter is a folder. If ommitted .lnk will be used by default.

 .Parameter Disable
  If used will create disabled Actions. Defaults to $false if ommitted (create Enabled Actions).

 .Example
   # Create WEM Actions from all the items in the default Start Menu locations.
   New-WEMActionXML

 .Example
   # Create WEM Actions from all the items in the default Start Menu locations and export this to a file that can be restored in WEM.
   New-WEMActionXML | Out-File $env:TEMP\VUEMApplications.xml

 .Example
   # Create WEM Actions from all the items in a custom folder, processing .exe and .lnk files.
   New-WEMActionXML -Path "E:\Custom Folder\Menu Items" -FileTypes exe,lnk

 .Example
   # Create WEM Actions from Notepad.exe
   New-WEMActionXML -Path "C:\Windows\System32\notepad.exe" -Name "Notepad example"
#>
function New-WEMActionXML {
    param(
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True)][string]$Path,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][bool] $Recurse = $true,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][string[]] $FileTypes,
        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)][switch]$Disable
    )

    # check if $Path is valid
    if ($Path -and !(Test-Path $Path)) {
         Write-Host "Cannot find path '$Path' because it does not exist." -ForegroundColor Red
         Break
    }

    # grab files
    $files = @()
    
    # tell user we're processing the Start Menu if $Path was not provided
    if (!$Path) {
        Write-Host "`nProcessing default Start Menu folders" -ForegroundColor Yellow
        $files = Get-FilesToProcess("$($env:ProgramData)\Microsoft\Windows\Start Menu\Programs")
        $files += Get-FilesToProcess("$($env:USERPROFILE)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs")
    } else {
        Write-Host "`nProcessing '$Path'" -ForegroundColor Yellow
        $files = Get-FilesToProcess($Path)
    }

    # pre-load System.Drawing namespace
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    # create a base XML file structure
    $StringWriter = New-Object System.IO.StringWriter
    $xmlWriter = New-Object System.Xml.XmlTextWriter $StringWriter
    $xmlWriter.Formatting = "indented"
    $xmlWriter.WriteProcessingInstruction("xml", "version='1.0' encoding='UTF-8'");
    
    # write xml root node
    $xmlWriter.WriteStartElement("ArrayOfVUEMApplication")
    $xmlWriter.WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
    $xmlWriter.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
    
    # process inputs
    ForEach($file in $files) {
        $Description = ""
        $WorkingDirectory = ""
        $IconLocation = ""
        $Arguments = ""
        $HotKey = "None"
        $TargetPath = ""
        $IconStream = ""
        if($file.Extension -eq ".lnk") {
            # grab lnk file for property parsing
            $obj = New-Object -ComObject WScript.Shell
            $lnk = $obj.CreateShortcut("$($file.FullName)")
            
            # grab properties
            $Description = $lnk.Description
            $WorkingDirectory = [System.Environment]::ExpandEnvironmentVariables($lnk.WorkingDirectory)
            $IconLocation = [System.Environment]::ExpandEnvironmentVariables($lnk.IconLocation.Split(",")[0])
            $Arguments = $lnk.Arguments
            if($lnk.Hotkey) { $HotKey = $lnk.Hotkey }
            $TargetPath = [System.Environment]::ExpandEnvironmentVariables($lnk.TargetPath)
            if(!$IconLocation) { $IconLocation = $TargetPath }
        } else {
            # for anything but .lnk try to generate relevant properties
            $WorkingDirectory = $file.FolderName
            $IconLocation = $file.FullName
            $TargetPath = $file.FullName
        }

        # only work if we have a target
        if($TargetPath) {
            # grab icon
            if($IconLocation -and (Test-Path $IconLocation)) { $IconStream = Get-IconToBase64 ([System.Drawing.Icon]::ExtractAssociatedIcon("$($IconLocation)")) }
            
            # write xml child node
            $xmlWriter.WriteStartElement("VUEMApplication")
            $xmlWriter.WriteElementString("State",$State)
            $xmlWriter.WriteElementString("IconIndex","0")
            $xmlWriter.WriteElementString("Name","$($file.BaseName)")
            $xmlWriter.WriteElementString("Description",$Description)
            $xmlWriter.WriteElementString("DisplayName",$file.BaseName)
            $xmlWriter.WriteElementString("StartMenuTarget","Start Menu\Programs$($file.RelativePath)")
            $xmlWriter.WriteElementString("TargetPath",$TargetPath)
            $xmlWriter.WriteElementString("Parameters",$Arguments)
            $xmlWriter.WriteElementString("WorkingDirectory",$WorkingDirectory)
            $xmlWriter.WriteElementString("WindowStyle","Normal")
            $xmlWriter.WriteElementString("IconLocation",$IconLocation)
            $xmlWriter.WriteElementString("Hotkey",$HotKey)
            $xmlWriter.WriteElementString("IconStream",$IconStream)
            $xmlWriter.WriteEndElement()
        }
    }

    # finish the XML Document
    $xmlWriter.WriteEndElement()
    $xmlWriter.Finalize
    $StringWriter.Flush();

    # return the raw xml data
    return $StringWriter.ToString();
}

<#
 .SYNOPSIS
  Helper function to grab an array of files to process.
#>
function Get-FilesToProcess{
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )

    if ((Get-Item $Path) -is [System.IO.DirectoryInfo]) {
        # what to do if $Path is a folder
        $relativePath = $Path
        if(!$Path.EndsWith("\*")) { 
            $Path += "\*"
        } else {
            $relativePath = $Path.Replace("\*","")
        }

        # build filter
        $Filter = "*.lnk"
        if($FileTypes) { 
            $Filter = @()
            foreach($FileType in ($FileTypes.Split(","))) {
                $Filter += "*." + $($FileType.Replace("*.","").Trim())
            }
        }

        # to recurse or not to recurse
        if($Recurse){
            $files = Get-ChildItem -Path $Path -Include $Filter -Recurse | Select-Object Name, @{ n = 'BaseName'; e = { $_.BaseName } }, @{ n = 'FolderName'; e = { Convert-Path $_.PSParentPath } }, @{ n = 'RelativePath'; e = { (Convert-Path $_.PSParentPath).Replace($relativePath,"") } }, FullName, @{ n = 'Extension'; e = { $_.Extension } } | Where-Object { !$_.PSIsContainer }
        } else {
            $files = Get-ChildItem -Path $Path -Include $Filter | Select-Object Name, @{ n = 'BaseName'; e = { $_.BaseName } }, @{ n = 'FolderName'; e = { Convert-Path $_.PSParentPath } }, @{ n = 'RelativePath'; e = { (Convert-Path $_.PSParentPath).Replace($relativePath,"") } }, FullName, @{ n = 'Extension'; e = { $_.Extension } } | Where-Object { !$_.PSIsContainer }
        }

    } else {
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
	return ([System.Convert]::ToBase64String($stream.ToArray()))
}

# expose the functions
Export-ModuleMember -function New-WEMActionXML

#New-WEMActionXML | Out-File d:\VUEMApplications.xml
#New-WEMActionXML -path 'C:\ProgramData\Microsoft\Windows\Start Menu' | Out-File d:\test.xml