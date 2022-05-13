#Requires -Version 3.0

[int]$AppV5MajorVersion = 1;
[int]$AppV5MinorVersion = 0;
[int]$AppV5BuildNumber = 5660;
[DateTime]$AppV5PublishDate = "2015-07-01T00:00:00";

Function Get-VEAppV5Version()
{
    <#
    .SYNOPSIS
    This function returns the version of the Virtual Engine AppV 5.0 PowerShell modules
    .DESCRIPTION
    This function returns the installed version of the Virtual Engine AppV 5.0 PowerShell modules
    .EXAMPLE
    Get-VEAppV5Version
    Returns the registered version information.
    .NOTES
    NAME: Get-VEAppV5Version
    AUTHOR: Iain Brighton, Virtual Engine
    LASTEDIT: 26/04/2013 16:35
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,VirtualEngine,AppV5
    .LINK
    http://virtualengine.co.uk
    #>

    Process
    {
        [string] $AppV5DisplayVersion = $AppV5MajorVersion.ToString() + "." + $AppV5MinorVersion.ToString();
        return [PSCustomObject]@{"DisplayVersion"=$AppV5DisplayVersion;"MajorVersion"=$AppV5MajorVersion;"MinorVersion"=$AppV5MinorVersion;BuildNumber=$AppV5BuildNumber;PublishDate=$AppV5PublishDate}
    }
}

Function Save-AppV5File()
{
    <#
    .SYNOPSIS
    This function extracts a file from an App-V 5.0 package
    .DESCRIPTION
    This function extracts the specified (case-sensitive) file from an App-V 5.0 compressed .APPV package/archive
    .EXAMPLE
    Save-AppV5File -AppV c:\package.APPV -File AppxManifest.xml -FilePath c:\temp\
    Extracts the AppxManifest.xml file from the App-V 5.0 c:\package.APPV file and saves it as c:\temp\appxmanifest.xml.
    .EXAMPLE
    Save-AppV5File -AppV c:\package.appv -File StreamMap.xml -FilePath c:\temp\ -Overwrite
    Extracts the StreamMap.xml file from the App-V 5.0 c:\package.appv file and saves it as c:\temp\streammap.xml and overwrites any existing file.
    .EXAMPLE
    Save-AppV5File -AppV c:\package.appv -File "Root/install.log" -FilePath c:\temp\
    Extracts the install.log file from the "Root/" directory within the c:\package.appv App-V 5.0 .APPV file, saving it as c:\temp\install.log.
    .PARAMETER AppV
    The source App-V 5.0 package/archive .APPV file
    .PARAMETER File
    The source CASE-SENSITIVE filename within the App-V 5.0 package/archive .APPV file to extract. Note: path names are specified as /
    .PARAMETER FilePath
    The target directory to save the extracted XML file to
    .PARAMETER Overwrite
    Whether to overwrite the existing file if it exists
    .NOTES
    NAME: Save-AppV5File
    AUTHOR: Iain Brighton, Virtual Engine
    LASTEDIT: 01/05/2013 15:51
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,AppxManifest.xml
    .LINK
    http://virtualengine.co.uk
    #>

    Param
    (
        [parameter(Mandatory=$true,Position=0,HelpMessage="Enter the source App-V 5.0 .APPV package/archive filename.")]
            [validatescript({
                if(!(Test-Path $_ -PathType Leaf)) {throw "$_ file not found";}
                else {return $true;}
                })]
            [alias("AppVFile","FullName")]
            [string]$AppV,
        [parameter(Mandatory=$true,Position=1,HelpMessage="Enter the CASE-SENSITIVE file to extract from the source .APPV file")]
            [alias("Source")]
            [string]$File,
        [parameter(Position=2,HelpMessage="Enter the target directory to save the extracted file to.")]
            [validatescript({
                if(!(Test-Path $_ -PathType Container)) {throw "$_ target path not found";}
                else {return $true;}
            })]
            [alias("SaveAs","Target")]
            [string]$FilePath,
        [Parameter(HelpMessage="Whether to overwrite an existing file.")]
            [switch]$Overwrite
    )

    Process
    {
        ### The System.IO.Compression.FileSystem requires at least .Net Framework 4.5
        Write-Verbose "Save-AppV5File: Loading .Net Framework assemblies";
        [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression") | Out-Null;
        [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null;

        ### Open the ZipArchive with read access
        Write-Verbose "Save-AppV5File: Opening .APPV archive $AppV";
        $AppV5Archive = New-Object System.IO.Compression.ZipArchive(New-Object System.IO.FileStream($AppV, [System.IO.FileMode]::Open));

        ### Locate the source XML file
        Write-Verbose "Save-AppV5File: Locating file $File within archive $AppV";
        $AppV5ArchiveEntry = $AppV5Archive.GetEntry($File);

        ### Check the source file is in the .APPV archive
        Write-Verbose "Save-AppV5File: Checking file object reference";
        if ($AppV5ArchiveEntry -eq $null)
        {
            ### Puppies will die!
            Write-Error "Save-AppV5File: The $File file does not exist within the source .APPV package or is not a valid .APPV package." -ForegroundColor Red;
            ### Ensure we close the file handle otherwise the file will be left open
            $AppV5Archive.Dispose();
            Write-Verbose "Save-AppV5File: Closing $AppV file handle";
            return $false;
        }

        ### Generate the -OutFile as required
        if (!($FilePath)) { $OutFile = Join-Path (Get-Item $AppV).DirectoryName $AppV5ArchiveEntry.Name; }
        else { $OutFile = Join-Path $FilePath $AppV5ArchiveEntry.Name; }
        Write-Verbose "Save-AppV5File: Saving file as: $OutFile";

        ### ZipArchiveEntry.ExtractToFile is an extension method
        Write-Verbose "Save-AppV5File: Extracting $File from $AppV";
        $OutputFile = [System.IO.Compression.ZipFileExtensions]::ExtractToFile($AppV5ArchiveEntry, $OutFile, $Overwrite);

        ### Ensure we close the file handle otherwise the file will be left open
        Write-Verbose "Save-AppV5File: Closing $AppV file handle";
        $AppV5Archive.Dispose();
        return (Get-Item $OutFile);
    }
}

Function Get-AppV5File()
{
   <#
    .SYNOPSIS
    This function extracts and returns the contents of an App-V 5.0 XML file
    .DESCRIPTION
    This function attempts to extract the contents of the specifed file from a compressed .APPV package/archive and return the contents as a System.Xml.XmlDocument object
    .EXAMPLE
    Get-AppV5File -AppV c:\package.appv -XML AppxManifest.xml
    Extracts the AppxManifest.xml file from the App-V 5.0 c:\package.appv file and returns the contents as a System.Xml.XmlDocument object.
    Note: For the standard .APPV XML files it is recommended you use the Get-AppV5FileXml (GAppV5FX) command.
    .PARAMETER AppV
    The source App-V 5.0 package/archive .APPV file
    .PARAMETER XML
    The CASE-SENSITIVE source XML file within the App-V 5.0 package/archive .APPV file to return
    .NOTES
    NAME: Get-AppV5File
    AUTHOR: Iain Brighton, Virtual Engine
    LASTEDIT: 26/04/2013 10:36
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,AppxManifest.xml,XmlDocument
    .LINK
    http://virtualengine.co.uk
    #>

    Param
    (
        [parameter(
                Mandatory=$true,
                Position=0,
                HelpMessage="Enter the source App-V 5.0 .APPV package/archive filename."
                )]
            [validatescript({
                if(!(Test-Path $_ -PathType Leaf)) {throw "$_ file not found";}
                else {return $true;}
                })]
            [alias("AppVFile","FullName")]
            [string]$AppV,
        [parameter(Mandatory=$true,Position=1,HelpMessage="Enter the CASE-SENSITIVE XML file within the .APPV file.")]
            [validatepattern("^*\.xml$")]
            [alias("Xml")]
            [string]$File
    )

    ### The System.IO.Compression.FileSystem requires at least .Net Framework 4.5
    Write-Verbose "Get-AppV5File: Loading .Net Framework assemblies";
    [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression") | Out-Null;
    [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null;

    ### Open the ZipArchive with read access
    Write-Verbose "Get-AppV5File: Opening .APPV archive $AppV";
    $AppV5Archive = New-Object System.IO.Compression.ZipArchive(New-Object System.IO.FileStream($AppV, [System.IO.FileMode]::Open));

    ### Locate the source XML file
    Write-Verbose "Get-AppV5File: Locating file $File within archive $AppV";
    $AppV5ArchiveEntry = $AppV5Archive.GetEntry($File);

    ### Check the source file is in the .APPV archive
    Write-Verbose "Get-AppV5File: Checking file object reference";
    if ($AppV5ArchiveEntry -eq $null)
    {
        ### Puppies will die!
        Write-Error "Get-AppV5File: The $File file does not exist within the source .APPV package" -ForegroundColor Red;
        ### Ensure we close the file handle otherwise the file will be left open
        Write-Verbose "Get-AppV5File: Closing $AppV file handle";
        $AppV5Archive.Dispose();
        return $null;
    }

    ### Create an new XmlDocument object
    Write-Verbose "Get-AppV5File: Creating new XmlDocument";
    $xmlDoc = New-Object System.Xml.XmlDocument;
    ### Load the XmlDocument using the System.IO.Stream from the ZipArchiveEntry object
    Write-Verbose "Get-AppV5File: Loading $File contents from $File into the XmlDocument";
    $xmlDoc.Load($AppV5ArchiveEntry.Open());

    ### Ensure we close the file handle otherwise the file will be left open
    Write-Verbose "Get-AppV5File: Closing $AppV file handle";
    $AppV5Archive.Dispose();

    ### Return our XmlDocument object
    return $xmlDoc;
}

Function Save-AppV5FileXml()
{
    <#
    .SYNOPSIS
    This function extracts an App-V 5.0 file
    .DESCRIPTION
    This function extracts a specific App-V 5.0 XML file from an App-V 5.0 compressed .APPV package/archive
    .EXAMPLE
    Save-AppV5FileXml -AppV c:\package.appv -XML AppxManifest -FilePath c:\temp\
    Extracts the AppxManifest.xml file from the App-V 5.0 c:\package.appv file and saves it as c:\temp\appxmanifest.xml.
    .EXAMPLE
    Save-AppV5FileXml -AppV c:\package.appv -XML StreamMap -FilePath c:\temp\ -Overwrite
    Extracts the StreamMap.xml file from the App-V 5.0 c:\package.appv file and saves it as c:\temp\streammap.xml, overwriting it if it exists.
    .PARAMETER AppV
    The source App-V 5.0 package/archive .APPV file
    .PARAMETER XML
    The specific XML file within the App-V 5.0 package/archive .APPV file to extract
    .PARAMETER FilePath
    The target directory to save the extracted XML file to
    .NOTES
    NAME: Save-AppV5FileXml
    AUTHOR: Iain Brighton, Virtual Engine
    LASTEDIT: 01/05/2013 15:56
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,.xml
    .LINK
    http://virtualengine.co.uk
    #>

    Param
    (
        [parameter(Mandatory=$true,Position=0,HelpMessage="Enter the source App-V 5.0 .APPV package/archive filename.")]
            [validatescript({
                if(!(Test-Path $_ -PathType Leaf)) {throw "$_ file not found";}
                else {return $true;}
            })]
            [alias("AppVFile","SourceFile","SourceAppVArchive","Archive")]
            [string]$AppV,
        [parameter(Mandatory=$true,Position=1,HelpMessage="Select the source .APPV XML file")]
            [ValidateSet("AppxManifest","AppxBlockMap","FilesystemMetadata","PackageHistory","StreamMap")]
            [alias("Source")]
            [string]$XML,
        [parameter(Position=2,HelpMessage="Enter the target directory to save the extracted file to.")]
            [validatescript({
                if(!(Test-Path $_ -PathType Container)) {throw "$_ target path not found";}
                else {return $true;}
            })]
            [alias("SaveAs","Target")]
            [string]$FilePath,
        [parameter(HelpMessage="Whether to overwrite an existing file.")]
            [switch]$Overwrite
    )

    Process
    {
        Write-Verbose "Save-AppV5FileXml: Locating XML file reference";
        switch ($XML.ToLower())
        {
            "appxmanifest"
                {
                    if ($FilePath) { return Save-AppV5File -AppV $AppV -File "AppxManifest.xml" -FilePath $FilePath $ -Overwrite; }
                    else { return Save-AppV5File -AppV $AppV -File "AppxManifest.xml" -Overwrite; }
                }
            "streammap"
                {
                    if ($FilePath) { return Save-AppV5File -AppV $AppV -File "StreamMap.xml" -FilePath $FilePath -Overwrite; }
                    else { return Save-AppV5File -AppV $AppV -File "StreamMap.xml" -Overwrite; }
                }
            "appxblockmap"
                {
                    if ($FilePath) { return Save-AppV5File -AppV $AppV -File "AppxBlockMap.xml" -FilePath $FilePath -Overwrite; }
                    else { return Save-AppV5File -AppV $AppV -File "SAppxBlockMap.xml" -Overwrite; }
                }
            "packagehistory"
                {
                    if ($FilePath) { return Save-AppV5File -AppV $AppV -File "PackageHistory.xml" -FilePath $FilePath -Overwrite; }
                    else { return Save-AppV5File -AppV $AppV -File "PackageHistory.xml" -Overwrite; }
                }
            "filesystemmetadata"
                {
                    if ($FilePath) { return Save-AppV5File -AppV $AppV -File "FilesystemMetadata.xml" -FilePath $FilePath -Overwrite; }
                    else { return Save-AppV5File -AppV $AppV -File "FilesystemMetadata.xml" -Overwrite; }
                }
            default { return $null; }
        }
    }
}

Function Get-AppV5FileXml()
{
    <#
    .SYNOPSIS
    This function extracts and returns the contents of the specified App-V 5.0 file
    .DESCRIPTION
    This function attempts to extract the contents of a specific App-V 5.0 XML file from a compressed .APPV package/archive and return the contents as a System.Xml.XmlDocument object. This function attempts to extract the contents of a specific App-V 5.0 XML file from a compressed .APPV package/archive and return the contents as a System.Xml.XmlDocument object. The -XML parameter has a specific set of options and is not case-sensitive (unlike the Get-AppV5File module).
    .EXAMPLE
    Get-AppV5FileXml -AppV c:\package.appv -XML AppxManifest
    Extracts the AppxManifest.xml file from the App-V 5.0 c:\package.appv file and returns the contents as an System.Xml.XmlDocument.
    .PARAMETER AppV
    The source App-V 5.0 package/archive .APPV file
    .PARAMETER XML
    The specific App-V 5.0 XML file within the App-V 5.0 package/archive .APPV file to extract
    .NOTES
    NAME: Get-AppV5FileXml
    AUTHOR: Iain Brighton, Virtual Engine
    LASTEDIT: 26/04/2013 12:15
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,.xml
    .LINK
    http://virtualengine.co.uk
    #>

    Param
    (
        [parameter(Mandatory=$true,Position=0,HelpMessage="Enter the source App-V 5.0 .APPV package/archive filename.")]
            [validatescript({
                if(!(Test-Path $_ -PathType Leaf)) {throw "$_ file not found";}
                else {return $true;}
            })]
            [alias("AppVFile","SourceFile","SourceAppVArchive","Archive")]
            [string]$AppV,
        [parameter(Mandatory=$true,Position=1,HelpMessage="Select the source .APPV XML file")]
            [validateset("AppxManifest","AppxBlockMap","FilesystemMetadata","PackageHistory","StreamMap")]
            [alias("Source")]
            [string]$XML
    )

    Process
    {
        Write-Verbose "Get-AppV5FileXml: Locating XML file reference";
        switch ($XML.ToLower())
        {
            "appxmanifest" { return Get-AppV5File -AppV $AppV -File "AppxManifest.xml"; }
            "streammap" { return Get-AppV5File -AppV $AppV -File "StreamMap.xml"; }
            "appxblockmap" { return Get-AppV5File -AppV $AppV -File "AppxBlockMap.xml"; }
            "packagehistory" { return Get-AppV5File -AppV $AppV -File "PackageHistory.xml"; }
            "filesystemmetadata" { return Get-AppV5File -AppV $AppV -File "FilesystemMetadata.xml"; }
            default { return $null; }
        }
    }
}

Function Save-AppV5FileXmlPackage()
{
    <#
    .SYNOPSIS
    This function saves App-V 5.0 .APPV package information to an XML file
    .DESCRIPTION
    This function saves an XML file containing information from the AppxManifest.xml, PackageHistory.xml, StreamMap.xml and FilesystemMetadata.xml files in the specified App-V 5.0 .APPV package.
    .EXAMPLE
    Save-AppV5FileXmlPackage -AppV c:\package.appv -FilePath c:\temp\
    Saves information the AppxManifest.xml, PackageHistory.xml, StreamMap.xml and FilesystemMetadata.xml files within an App-V 5.0 .APPV package as a single c:\temp\AppV5Package.xml file.
    .PARAMETER AppV
    The source App-V 5.0 package/archive .APPV file
    .PARAMETER FilePath
    The target directory to save the generated AppV5Package.xml file to
    .NOTES
    NAME: Save-AppV5FileXmlPackage
    AUTHOR: Iain Brighton, Virtual Engine
    LASTEDIT: 01/04/2013 16:01
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,AppxManifest.xml,XmlDocument
    .LINK
    http://virtualengine.co.uk
    #>

    Param
    (
        [parameter(Mandatory=$true,Position=0,HelpMessage="Enter the source App-V 5.0 .APPV package/archive filename.")]
            [validatescript({
                if(!(Test-Path $_ -PathType Leaf)) {throw "$_ file not found";}
                else {return $true;}
            })]
            [alias("AppVFile","SourceFile","SourceAppVArchive","Archive")]
            [string]$AppV,
        [parameter(Position=1,HelpMessage="Enter the full path\filename to save the generated XML file as.")]
            [validatescript({
                if(!(Test-Path $_ -PathType Container)) {throw "$_ target path not found";}
                else {return $true;}
            })]
            [alias("SaveAs","Target")]
            [string]$FilePath
    )

    Process
    {
         ### Generate the -OutFile as required
        if (!($FilePath)) { $OutFile = Join-Path (Get-Item $AppV).DirectoryName ((Get-Item $AppV).BaseName + "_Package.xml"); }
        else { $OutFile = Join-Path $FilePath ((Get-Item $AppV).BaseName + "_Package.xml"); }
        Write-Verbose "Save-AppV5FileXmlPackage: Saving file as: $OutFile";
        $AppV5FilePackage = (Get-AppV5FileXmlPackage -AppV $AppV).Save($OutFile);
        return (Get-Item $OutFile);;
    }
}

Function Get-AppV5FileXmlPackage()
{
    <#
    .SYNOPSIS
    This function returns App-V 5.0 .APPV package information in XML
    .DESCRIPTION
    This function returns a System.Xml.XmlDocument containing information from the AppxManifest.xml, PackageHistory.xml, StreamMap.xml and FilesystemMetadata.xml files in the specified App-V 5.0 .APPV package. Data is returned under the <AppV5> element/node.
    .EXAMPLE
    Get-AppV5FileXmlPackage -AppV c:\package.appv
    Returns details of the c:\package.appv App-V 5.0 package .APPV files as a System.Xml.XmlDocument
    .PARAMETER AppV
    The source App-V 5.0 package/archive .APPV file
    .NOTES
    NAME: Get-AppV5FileXmlPackage
    AUTHOR: Iain Brighton, Virtual Engine
    LASTEDIT: 26/04/2013 13:20
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,AppxManifest.xml,XmlDocument
    .LINK
    http://virtualengine.co.uk
    #>

    Param
    (
        [parameter(Mandatory=$true,Position=0,HelpMessage="Enter the source App-V 5.0 .APPV package/archive filename.")]
            [validatescript({
                if(!(Test-Path $_ -PathType Leaf)) {throw "$_ file not found";}
                else {return $true;}
            })]
            [alias("AppVFile","SourceFile","SourceAppVArchive","Archive")]
            [string]$AppV
    )

    Process
    {
        ### Create the base System.Xml.XmlDocument
        Write-Verbose "Get-AppV5FileXmlPackage: Creating new XmlDocument";
        [System.Xml.XmlDocument]$xmlDoc = New-Object System.Xml.XmlDocument;

        ### Create the <Appv5> document element
        Write-Verbose "Get-AppV5FileXmlPackage: Creating <AppV5> root element";
        $xmlRoot = ($xmlDoc.AppendChild($xmlDoc.CreateElement("AppV5")));

        Write-Verbose "Get-AppV5FileXmlPackage: Loading AppxManifest.xml from $AppV";
        $xmlRoot.AppendChild(($xmlDoc.ImportNode((Get-AppV5FileXml -AppV $AppV -XML AppxManifest).DocumentElement, $true))) | Out-Null;
        Write-Verbose "Get-AppV5FileXmlPackage: Loading PackageHistory.xml from $AppV";
        $xmlRoot.AppendChild(($xmlDoc.ImportNode((Get-AppV5FileXml -AppV $AppV -XML PackageHistory).DocumentElement, $true))) | Out-Null;
        Write-Verbose "Get-AppV5FileXmlPackage: Loading StreamMap.xml from $AppV";
        $xmlRoot.AppendChild(($xmlDoc.ImportNode((Get-AppV5FileXml -AppV $AppV -XML StreamMap).DocumentElement, $true))) | Out-Null;
        Write-Verbose "Get-AppV5FileXmlPackage: Loading FilesystemMetadata.xml from $AppV";
        $xmlRoot.AppendChild(($xmlDoc.ImportNode((Get-AppV5FileXml -AppV $AppV -XML FilesystemMetadata).DocumentElement, $true))) | Out-Null;
        Write-Verbose "Get-AppV5FileXmlPackage: Returning XmlDocument";
        return $xmlDoc;
    }
}

Function Get-AppV5FilePackage()
{
    <#
    .SYNOPSIS
    This function returns various properties from an App-V 5.0 .APPV package as an object
    .DESCRIPTION
    This function returns a custom PowerShell object containing pertinent information within the AppxManifest.xml, PackageHistory.xml, StreamMap.xml and FilesystemMetadata.xml files in the specified App-V 5.0 .APPV package.
    .EXAMPLE
    Get-AppV5FilePackage -AppV c:\package.appv
    Returns details of the c:\package.APPV App-V 5.0 package .APPV files as a custom PowerShell object.
    .PARAMETER AppV
    The source App-V 5.0 package/archive .APPV file
    .NOTES
    NAME: Get-AppV5FilePackage
    AUTHOR: Iain Brighton, Virtual Engine
    LASTEDIT: 26/04/2013 13:20
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,AppxManifest.xml,XmlDocument
    .LINK
    http://virtualengine.co.uk
    #>

    Param
    (
        [parameter(Mandatory=$true,Position=0,HelpMessage="Enter the source App-V 5.0 .APPV package/archive filename.")]
            [validatescript({
                if(!(Test-Path $_ -PathType Leaf)) {throw "$_ file not found";}
                else {return $true;}
            })]
            [alias("File","AppVFile","SourceFile","SourceAppVArchive","Archive")]
            [string]$AppV
    )

    Process
    {
        ### Process the file details
        $FileInfo = Get-Item $AppV;

        ### Create the custom AppV5File object
        ### The following properties are available in the AppxManifest.xml file
        $AppV5FilePackage = New-Object PSObject;
        $AppV5FilePackage | Add-Member -Name Name -Value $FileInfo.Name -MemberType NoteProperty;
        $AppV5FilePackage | Add-Member -Name PackageId -Value "" -MemberType NoteProperty;
        $AppV5FilePackage | Add-Member -Name VersionId -Value "" -MemberType NoteProperty;
        $AppV5FilePackage | Add-Member -Name DisplayName -Value "" -MemberType NoteProperty;
        $AppV5FilePackage | Add-Member -Name AppVPackageDescription -Value "" -MemberType NoteProperty;
        $AppV5FilePackage | Add-Member -Name Version -Value "" -MemberType NoteProperty;
        $AppV5FilePackage | Add-Member -Name OSMinVersion -Value "" -MemberType NoteProperty;
        $AppV5FilePackage | Add-Member -Name OSMaxVersionTested -Value "" -MemberType NoteProperty;
        $AppV5FilePackage | Add-Member -Name SequencingStationProcessorArchitecture -Value "" -MemberType NoteProperty;
        ### The following properties are calculated from the .APPV archive
        $AppV5FilePackage | Add-Member -Name UncompressedSize -Value 0 -MemberType NoteProperty;
        ### The following properties are available in the StreamMap.xml
        $AppV5FilePackage | Add-Member -Name PrimaryFeatureBlockLoadAll -Value "" -MemberType NoteProperty;
        ### The following properties are available in the FilesystemMetadata.xml
        $AppV5FilePackage | Add-Member -Name FileSystemRoot -Value "" -MemberType NoteProperty;
        $AppV5FilePackage | Add-Member -Name FileSystemShort -Value "" -MemberType NoteProperty;

        ### Add the FileInfo
        $AppV5FilePackage | Add-Member -Name FileInfo -Value $FileInfo -MemberType NoteProperty;


        ### Retrieve the AppV5FilePackage
        $AppV5FileXmlPackage = Get-AppV5FileXmlPackage -AppV $AppV;

        ### Retrieve and assign the AppV5FilePackage properties/details
        $AppV5FilePackage.PackageId = $AppV5FileXmlPackage.AppV5.Package.Identity.PackageId;
        $AppV5FilePackage.VersionId = $AppV5FileXmlPackage.AppV5.Package.Identity.VersionId;
        $AppV5FilePackage.DisplayName = $AppV5FileXmlPackage.AppV5.Package.Properties.DisplayName;
        $AppV5FilePackage.AppVPackageDescription = $AppV5FileXmlPackage.AppV5.Package.Properties.AppVPackageDescription;
        $AppV5FilePackage.Version = $AppV5FileXmlPackage.AppV5.Package.Identity.Version;
        $AppV5FilePackage.OSMinVersion = $AppV5FileXmlPackage.AppV5.Package.Prerequisites.OSMinVersion;
        $AppV5FilePackage.OSMaxVersionTested = $AppV5FileXmlPackage.AppV5.Package.Prerequisites.OSMaxVersionTested;
        $AppV5FilePackage.SequencingStationProcessorArchitecture = $AppV5FileXmlPackage.AppV5.Package.Prerequisites.TargetOSes.SequencingStationProcessorArchitecture;

        ### Add the Asset Intelligence properties
        $AppV5FileAssets = New-Object System.Collections.ArrayList;
        foreach ($Asset in $AppV5FileXmlPackage.AppV5.Package.AssetIntelligence.ChildNodes)
        {
            $AppV5FilePackageAsset = New-Object PSObject;
            $AppV5FilePackageAsset | Add-Member -Name SoftwareCode -Value $Asset.SoftwareCode -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name ProductName -Value $Asset.ProductName -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name ProductVersion -Value $Asset.ProductVersion -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name Publisher -Value $Asset.Publisher -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name ProductID -Value $Asset.ProductID -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name Language -Value $Asset.Language -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name ChannelCode -Value $Asset.ChannelCode -MemberType NoteProperty;
            if ($Asset.InstallDate -ne "") {
                try {
                    ## Fix bug reported by J. van Gessel parsing 'maandag 15 juni 2015 15:55:54' and 'Wed Jan 21 09:15:11 GMT 2015'
                    $AppV5FilePackageAsset | Add-Member -Name InstallDate -Value ([datetime]::ParseExact($Asset.InstallDate, 'yyyyMMdd', $null)) -MemberType NoteProperty;
                }
                catch { }
            }
            $AppV5FilePackageAsset | Add-Member -Name RegisteredUser -Value $Asset.RegisteredUser -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name InstalledLocation -Value $Asset.InstalledLocation -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name CM_DSLID -Value $Asset.CM_DSLID -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name VersionMajor -Value $Asset.VersionMajor -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name VersionMinor -Value $Asset.VersionMinor -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name ServicePack -Value $Asset.ServicePack -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name UpgradeCode -Value $Asset.UpgradeCode -MemberType NoteProperty;
            $AppV5FilePackageAsset | Add-Member -Name OsComponent -Value $Asset.OsComponent -MemberType NoteProperty;
            $AppV5FileAssets.Add($AppV5FilePackageAsset) | Out-Null;
        }
        ### Add the Asset Intelligence Collection/Array List to the AppV5FileDetails custom object
        $AppV5FilePackage | Add-Member -Name "AssetIntelligence" -MemberType NoteProperty -Value $AppV5FileAssets | Out-Null;

        ### Create the Application Collection/Array List
        $AppV5FilePackageApplications = New-Object System.Collections.ArrayList;
        foreach ($PackageApplication in $AppV5FileXmlPackage.AppV5.Package.Applications.ChildNodes)
        {
            ### Create the Application customer PSObejct
            $Application = New-Object PSObject;
            $Application | Add-Member -Name Id -Value $PackageApplication.Id -MemberType NoteProperty;
            $Application | Add-Member -Name Origin -Value $PackageApplication.Origin -MemberType NoteProperty;
            $Application | Add-Member -Name TargetInPackage -Value $PackageApplication.TargetInPackage -MemberType NoteProperty;
            $Application | Add-Member -Name Target -Value $PackageApplication.Target -MemberType NoteProperty;
            $Application | Add-Member -Name Name -Value $PackageApplication.VisualElements.Name -MemberType NoteProperty;
            $Application | Add-Member -Name Version -Value $PackageApplication.VisualElements.Version -MemberType NoteProperty;
            $AppV5FilePackageApplications.Add($Application) | Out-Null;
        }
        ### Add the Applications Collection/Array List to the AppV5FileDetails custom object
        $AppV5FilePackage | Add-Member -Name "Applications" -MemberType NoteProperty -Value $AppV5FilePackageApplications | Out-Null;

        ### Create the Package History Items collection
        $AppV5FilePackagePackageHistory = New-Object System.Collections.ArrayList;
        foreach ($PackageHistoryItem in $AppV5FileXmlPackage.AppV5.PackageHistory.ChildNodes)
        {
            ### Create the PackageHistoryItem custom PSObject
            $PackageHistory = New-Object PSObject;
            $PackageHistory | Add-Member -Name Time -Value ([DateTime]$PackageHistoryItem.Time) -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name PackageVersion -Value $PackageHistoryItem.PackageVersion -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name SequencerVersion -Value $PackageHistoryItem.SequencerVersion -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name SequencerUser -Value $PackageHistoryItem.SequencerUser -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name SequencingStation -Value $PackageHistoryItem.SequencingStation -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name WindowsVersion -Value $PackageHistoryItem.WindowsVersion -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name WindowsFolder -Value $PackageHistoryItem.WindowsFolder -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name UserFolder -Value $PackageHistoryItem.UserFolder -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name SystemType -Value $PackageHistoryItem.SystemType -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name Processor -Value $PackageHistoryItem.Processor -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name LastRebootNormal -Value $PackageHistoryItem.LastRebootNormal -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name TerminalServices -Value $PackageHistoryItem.TerminalServices -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name RemoteSession -Value $PackageHistoryItem.RemoteSession -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name NetFrameworkVersion -Value $PackageHistoryItem.NetFrameworkVersion -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name IEVersion -Value $PackageHistoryItem.IEVersion -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name PackageOSBitness -Value $PackageHistoryItem.PackageOSBitness -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name PackagingEngine -Value $PackageHistoryItem.PackagingEngine -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name Locale -Value $PackageHistoryItem.Locale -MemberType NoteProperty;
            $PackageHistory | Add-Member -Name InUpgrade -Value $PackageHistoryItem.InUpgrade -MemberType NoteProperty;
            $AppV5FilePackagePackageHistory.Add($PackageHistory) | Out-Null;
        }
        ### Add the Package History Items Collection/Array List to the AppV5FileDetails custom object
        $AppV5FilePackage | Add-Member -Name "PackageHistory" -MemberType NoteProperty -Value $AppV5FilePackagePackageHistory | Out-Null;

        ### Retrieve the PrimaryFeatureBlock information
        if ($AppV5FileXmlPackage.AppV5.StreamMap.FeatureBlock[0].Id -eq "PrimaryFeatureBlock") { $AppV5FilePackage.PrimaryFeatureBlockLoadAll = $AppV5FileXmlPackage.AppV5.StreamMap.FeatureBlock[0].LoadAll; }
        else { $AppV5FilePackage.PrimaryFeatureBlockLoadAll = $AppV5FileXmlPackage.AppV5.StreamMap.FeatureBlock[1].LoadAll; }

        ### Retrieve the PVAD directory information
        $AppV5FilePackage.FileSystemRoot = $AppV5FileXmlPackage.AppV5.Metadata.Filesystem.Root;
        $AppV5FilePackage.FileSystemShort = $AppV5FileXmlPackage.AppV5.Metadata.Filesystem.Short;

        ### The System.IO.Compression.FileSystem requires at least .Net Framework 4.5
        Write-Verbose "Get-AppV5File: Loading .Net Framework assemblies";
        [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression") | Out-Null;
        [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null;

        ### Open the ZipArchive with read access
        Write-Verbose "Get-AppV5FilePackage: Opening .APPV archive $AppV";
        $AppV5Archive = New-Object System.IO.Compression.ZipArchive(New-Object System.IO.FileStream($AppV, [System.IO.FileMode]::Open));

        ### Create the Files collection/array list
        $AppV5ArchiveFiles = New-Object System.Collections.ArrayList;

        ### Total the compressed/uncompressed size of all files within the .APPV file
        foreach ($ZipArchiveEntry in $AppV5Archive.Entries)
        {
            $AppV5FilePackage.UncompressedSize += $ZipArchiveEntry.Length;
            ### Add the ZipArchiveEntry to the collection/array list
            $AppV5ArchiveFiles.Add($ZipArchiveEntry) | Out-Null;
        }
        ### Add the ArrayList of ZipArchiveEntry objects to the AppV5ArchiveDetails custom PSObject
        $AppV5FilePackage | Add-Member -Name "Files" -MemberType NoteProperty -Value $AppV5ArchiveFiles;

        ### Close the .APPV file handle
        $AppV5Archive.Dispose();

        return $AppV5FilePackage;
    }
}

Function Save-AppV5FileReport()
{
    <#
    .SYNOPSIS
    This function saves a HTML report of an App-V 5.0 .APPV package
    .DESCRIPTION
    This function saves an HTML report of the specified App-V 5.0 .APPV package.
    .EXAMPLE
    Save-AppV5FileReport -AppV c:\package.appv
    Creates an AppV5Report.html summary HTML report in the location of the c:\package.APPV App-V 5.0 .APPV package file.
    .EXAMPLE
    Save-AppV5FileReport -AppV c:\package.appv -FilePath c:\temp\ -Detailed
    Creates a detailed HTML report of the c:\package.APPV App-V 5.0 .APPV package file in the c:\temp\ directory.
    .EXAMPLE
    Save-AppV5FileReport -AppV c:\package.appv -Open
    Creates an AppV5Report.html summary HTML report in the location of the c:\package.APPV App-V 5.0 .APPV package file and opens it in the default browser.
    .EXAMPLE
    Save-AppV5FileReport -AppV c:\package.appv -FilePath c:\temp\ -CSS $CSS
    Creates a c:\temp\AppV5Report.html summary HTML report of the c:\package.appv App-V 5.0 .APPV package file using custom CSS contained in the $CSS variable.
    .PARAMETER AppV
    The source App-V 5.0 package/archive .APPV file
    .PARAMETER FilePath
    The target directory to save the generated AppV5Report.html file
    .PARAMETER CSS
    Custom Cascading Stlye Sheet (CSS), i.e. <style></style>
    .PARAMETER Detailed
    Create a detailed report including Asset Intelligence and the Package Contents
    .PARAMETER Open
    Opens the resulting HTML file in the default browser
    .NOTES
    NAME: Get-AppV5FileReport
    AUTHOR: Iain Brighton, Virtual Engine
    LASTEDIT: 28/04/2013 19:59
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,HTML,report
    .LINK
    http://virtualengine.co.uk
    #>

    Param
    (
        [parameter(Mandatory=$true,Position=0,HelpMessage="Enter the source App-V 5.0 .APPV package/archive filename.")]
            [validatescript({
                if(!(Test-Path $_ -PathType Leaf)) {throw "$_ file not found";}
                else {return $true;}
            })]
            [alias("File","AppVFile","SourceFile","SourceAppVArchive","Archive")]
            [string]$AppV,
        [parameter(Position=1,HelpMessage="Target .html filename")]
            [validatescript({
                if(!(Test-Path $_ -PathType Container)) {throw "$_ target path not found";}
                else {return $true;}
            })]
            [string]$FilePath,
        [parameter(Position=2,HelpMessage="Enter custom <Style></Style> CSS.")]
            [string]$CSS,
        [parameter(Position=3,HelpMessage="Create a detailed package report")]
            [switch]$Detailed,
        [parameter(Position=4,HelpMessage="Open the HTML report after saving")]
            [switch]$Open
    )

    Process
    {
        if ($Detailed) { $AppV5HTMLReport = Get-AppV5FileReport -AppV $AppV -CSS $CSS -Detailed; }
        else { $AppV5HTMLReport = Get-AppV5FileReport -AppV $AppV -CSS $CSS; }

        ### Generate the target file name as required
        if (!($FilePath)) { $OutFile = Join-Path (Get-Item $AppV).DirectoryName ((Get-Item $AppV).BaseName + "_Report.html"); }
        else { $OutFile = Join-Path $FilePath ((Get-Item $AppV).BaseName + "_Report.html"); }
        Write-Verbose "Save-AppV5FileReport: Saving file as: $OutFile";

        $AppV5HTMLReport | Out-File $OutFile;
        If ($Open) { Invoke-Item $OutFile; }
        return (Get-Item $OutFile);
    }
}

Function Get-AppV5FileReport()
{
    <#
    .SYNOPSIS
    This function generates a HTML report of an App-V 5.0 .APPV package
    .DESCRIPTION
    This function returns an HTML report of the specified App-V 5.0 .APPV package.
    .EXAMPLE
    Get-AppV5FileReport -AppV c:\package.appv
    Returns a summary HTML report of the c:\package.appv App-V 5.0 .APPV package file.
    .EXAMPLE
    Get-AppV5FileReport -AppV c:\package.appv -Detailed
    Returns a detailed HTML report of the c:\package.appc App-V 5.0 .APPV package file.
    .EXAMPLE
    Get-AppV5FileReport -AppV c:\package.appv -CSS $CSS
    Returns a summary HTML report of the c:\package.appv App-V 5.0 .APPV package file using custom CSS contained in the $CSS variable.
    .PARAMETER AppV
    The source App-V 5.0 package/archive .APPV file
    .PARAMETER CSS
    Custom Cascading Stlye Sheet (CSS), i.e. <style></style>
    .PARAMETER Detailed
    Create a detailed report including Asset Intelligence and the Package Contents
    .NOTES
    NAME: Get-AppV5FileReport
    AUTHOR: Iain Brighton, Virtual Engine
    LASTEDIT: 28/04/2013 19:59
    WEBSITE: http://www.virtualengine.co.uk
    KEYWORDS: App-V,App-V 5,.APPV,HTML,report
    .LINK
    http://virtualengine.co.uk
    #>

    Param
    (
        [parameter(Mandatory=$true,Position=0,HelpMessage="Enter the source App-V 5.0 .APPV package/archive filename.")]
            [ValidateScript({ Test-Path $_})]
            [alias("File","AppVFile","SourceFile","SourceAppVArchive","Archive")]
            [string]$AppV,
        [parameter(Position=1,HelpMessage="Enter custom <Style></Style> CSS.")]
            [string]$CSS,
        [parameter(Position=2,HelpMessage="Create a detailed package report")]
            [switch]$Detailed
    )

    Process
    {
        ### Define the default CSS for the HTML report. This can be overridden via the -CSS parameter
        $DefaultCss = @"
        <link href="http://fonts.googleapis.com/css?family=Yanone+Kaffeesatz" rel="stylesheet" type="text/css">
        <style type="text/css">
        <!--
        body { font-family: "Segoe UI", "Lucida Grande", "Helvetica Neue", sans-serif; font-size: 11pt; margin-left: 30px; }

        #report { width: 100%; }
        #logo { height: 67px; vertical-align: text-bottom; padding-left: 128px; background-image: url(http://virtualengine.co.uk/wp-content/uploads/VESignatureLogo.png); background-repeat: no-repeat; }

        h1, h2, h3, h4, h5, h6 { font-family: "Yanone Kaffeesatz", "Segoe UI", sans-serif; font-weight: normal; color: #04499d; margin: 16pt 0 10px 0; }
        h1 { font-size: 32px; }
        h2 { font-size: 24px; }
        h3 { font-size: 20px; }
        table{ border-collapse: collapse; border: none; color: black; margin-bottom: 11pt; }
        table td{ font-size: 11pt; padding-left: 0px; padding-right: 20px; text-align: left; }
        table th { font-size: 11pt; font-weight: bold; padding-left: 0px; padding-right: 20px; text-align: left; }
        table.list{ float: left; }
        table.list td:nth-child(1){ font-weight: bold; border-right: 1px grey solid; text-align: right; }
        table.list td:nth-child(2){ padding-left: 7px; }
        table tr:nth-child(even) { background: #f2f2f2; }
        div.column { }
        div.first{ padding-right: 50px; border-right: 1px  grey solid; }
        div.second{ padding-left: 30px; }
        #aslist td:first-child { font-weight: bold }
        -->
        </style>
"@

        ### Retrieve the App-V 5 package details

        $Package = Get-AppV5FilePackage -AppV $AppV;

        ### Generate the report title
        $ReportTitle = "App-V 5 Package Report: " + $Package.DisplayName;
        ### Generate the report timestamp
        $GenerationTime = (Get-Date).ToString("g");
        ### Generate the report heading
        $ReportHeading = "<div id=""logo""><h1>" + $ReportTitle + "</h1></div><span>Report Creation Time: $GenerationTime</span>";
        ### Has a custom CSS been provided?
        if ($CSS) { Write-Verbose "Get-AppV5FileReport: Using custom CSS"; $ReportCSS = $CSS; } else { Write-Verbose "Get-AppV5FileReport: Using default  CSS"; $ReportCSS = $DefaultCss; }
        ### Create the report <HTML><HEAD> block
        $ReportHeader = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd"><html><head><title>';
        $ReportHeader += $ReportTitle +'</title>' + $ReportCSS + '</head><body><div id="report">';
        ### Create the report </BODY></HTML> footer
        $ReportFooter = '</div></body></html>'

        ### Create the "Package Information/File Information Table elements
        $ReportTableTop = "<table><tr><td>"
        $ReportTableMiddle = "</td><td>"
        $ReportTableBottom = "</td></tr></table>"

        ### Create the "Package Information" report section
        $PackageInfo = $Package | Select-Object @{Name="Package Id";Expression={$_.PackageId}},
            @{Name="Version Id";Expression={$_.VersionId}},@{Name="Package Description";Expression={$_.AppVPackageDescription}},
            Version,@{Name="PVAD";Expression={$_.FileSystemRoot}},@{Name="PVAD 8.3";Expression={$_.FileSystemShort}},
            OSMinVersion,OSMaxVersionTested,@{Name="Sequencer Architecture";Expression={$_.SequencingStationProcessorArchitecture}};
        ### Convert PackageInfo into a HTML fragment
        $PackageDetails = $PackageInfo | ConvertTo-Html -PreContent "<div class=""first column""><h2>Package Information</h2><div id=""aslist"">" -PostContent "</div></div>" -As List -Fragment | Out-String;

        ### Create the "File Information" report section
        $FileInformation = $Package.FileInfo;
        ### Add the UncompressedSize and Files.Count properties from the $Package.FileInfo property
        $FileInformation | Add-Member -Name "PrimaryFeatureBlockLoadAll" -Value $Package.PrimaryFeatureBlockLoadAll -MemberType NoteProperty;
        $FileInformation | Add-Member -Name "UncompressedSize" -Value $Package.UncompressedSize -MemberType NoteProperty;
        $FileInformation | Add-Member -Name "Files" -Value $Package.Files.Count -MemberType NoteProperty;
        ### Convert the FileInformation object into a HTML fragment
        $FileDetails = $FileInformation |
            Select-Object Name,@{Name="Directory Name";Expression={$_.DirectoryName}},@{Name="Load All";Expression={$_.PrimaryFeatureBlockLoadAll}},
                @{Name="Uncompressed Size";Expression={($_.UncompressedSize/1MB).ToString("N1") + " MB"}},
                @{Name="Compressed Size";Expression={($_.Length/1MB).ToString("N1") + " MB"}},@{Name="Files";Expression={$_.Files.ToString("N0")}},
                @{Name="Creation Time";Expression={$_.CreationTime}},@{Name="Last Access Time";Expression={$_.LastAccessTime}},
                @{Name="Last Write Time";Expression={$_.LastWriteTime}} |
            ConvertTo-Html -PreContent " <div class=""second column""><h2>File Information</h2><div id=""aslist"">" -PostContent "</div></div>" -As List -Fragment | Out-String;

        ### Create the "Asset Intelligence" report section and create the HTML fragment
        $AssetIntelligence = $Package.AssetIntelligence |
            Select-Object @{Name="Software Code";Expression={$_.SoftwareCode}},@{Name="Product Name";Expression={$_.ProductName}},
            @{Name="Product Version";Expression={$_.ProductVersion}},Publisher,Language,
            @{Name="Install Date";Expression={$_.InstallDate.ToString("d")}},@{Name="Major Version";Expression={$_.VersionMajor}},
            @{Name="Minor Version";Expression={$_.VersionMinor}} |
            ConvertTo-Html -PreContent "<h2>Asset Intelligence</h2>" -Fragment | Out-String;

        ### Create the "Applications" report section and generate the HTML fragment
        $Applications = $Package.Applications |
            Select-Object @{Name="Application Id";Expression={$_.Id}},Name,Target,Version,@{Name="In Package";Expression={$_.TargetInPackage}} |
            ConvertTo-Html -PreContent "<h2>Applications</h2>" -Fragment | Out-String;

        ### Create the "Package History" report section and retrive the HTML fragment
        $PackageHistory = $Package.PackageHistory | Select-Object Time,@{Name="Package Version";Expression={$_.PackageVersion}},
            @{Name="Sequencer Version";Expression={$_.SequencerVersion}},@{Name="Sequencer";Expression={$_.SequencingStation}},
            @{Name="Windows Version";Expression={$_.WindowsVersion}},@{Name=".Net Version";Expression={$_.NetFrameworkVersion}},
            @{Name="IE Version";Expression={$_.IEVersion}},@{Name="OS";Expression={$_.PackageOSBitness}},Locale |
            ConvertTo-Html -PreContent "<h2>Package History</h2>" -Fragment |
            Out-String

        ### Create the "Package Contents" report section and get the HTML fragment
        $PackageFiles = $Package.Files |
            Select-Object Name,@{Name="Full Name";Expression={$_.FullName}},@{Name="Last Write";Expression={$_.LastWriteTime.ToString("g")}},
                @{Name="File Size";Expression={($_.Length/1KB).ToString("N1") + " KB"}},
                @{Name="Compressed";Expression={($_.CompressedLength/1KB).ToString("N1") + " KB"}} |
            ConvertTo-Html -PreContent "<h2>Package Contents</h2>" -Fragment |
            Out-String;

        ### Piece together the entire HTML report
        $HTML = $ReportHeader + $ReportHeading;
        $HTML += $ReportTableTop + $PackageDetails + $ReportTableMiddle + $FileDetails + $ReportTableBottom;
        $HTML += $Applications + $PackageHistory;
        if ($Detailed) { $HTML += $AssetIntelligence + $PackageFiles; }
        $HTML += "<div style=""color:#bbbbbb;font-size:10pt;"">Report generated by the <a href=""http://virtualengine.co.uk/appv/"" style=""text-decoration: none"">App-V 5.0 .APPV PowerShell CmdLets</a> from <a href=""http://virtualegnine.co.uk"" style=""text-decoration: none"">Virtual Engine</a></div>" + $ReportFooter;

        ### Return the full HTML representation
        return $HTML;
    }
}

New-Alias -Name GAppV5F -Value Get-AppV5File -Description "Alias for the Get-AppV5File module";
New-Alias -Name SAppV5F -Value Save-AppV5File -Description "Alias for the Save-AppV5File module";
New-Alias -Name SAppV5FX -Value Save-AppV5FileXml -Description "Alias for the Save-AppV5FileXml module";
New-Alias -Name GAppV5FX -Value Get-AppV5FileXml -Description "Alias for the Get-AppV5FileXml module";
New-Alias -Name SAppV5FXP -Value Save-AppV5FileXmlPackage -Description "Alias for the Save-AppV5FileXmlPackage module";
New-Alias -Name GAppV5FXP -Value Get-AppV5FileXmlPackage -Description "Alias for the Get-AppV5FileXmlPackage module";
New-Alias -Name GAppV5FP -Value Get-AppV5FilePackage -Description "Alias for the Get-AppV5FilePackage module";
New-Alias -Name SAppV5FR -Value Save-AppV5FileReport -Description "Alias for the Save-AppV5FileReport module";
New-Alias -Name GAppV5FR -Value Get-AppV5FileReport -Description "Alias for the Get-AppV5FileReport module";

Export-ModuleMember -Alias * -Function *;

# SIG # Begin signature block
# MIIcYQYJKoZIhvcNAQcCoIIcUjCCHE4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUY9MnV5oyipg9ONPgUq2/0N5K
# nNOggheQMIIFGTCCBAGgAwIBAgIQA1YkzuBwY6CTUsB/f/3MCTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE1MDUxOTAwMDAwMFoXDTE3MDgy
# MzEyMDAwMFowYDELMAkGA1UEBhMCR0IxDzANBgNVBAcTBk94Zm9yZDEfMB0GA1UE
# ChMWVmlydHVhbCBFbmdpbmUgTGltaXRlZDEfMB0GA1UEAxMWVmlydHVhbCBFbmdp
# bmUgTGltaXRlZDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKi0Jmm3
# YpnELWD00PUUo7904RJhU0SbfzM7HmGLWaQU3EegZ0n4/EX5AuGhmb7KRYB9HBj/
# 0jSYdcOu24l/aE+TRS/bFCR1LFXcjwQNS4aVnBdPTz9vJhhijt6EByOIUyw/HrYi
# JRfahWp7XdT5qNI3ak+LhK0psdJuJGBlFE2uWTwCtFGUrgIjV+ojSC/q4/wT2yN5
# 7jyc88eN/KB2JGRkbbMddfqSwOoSCslM0biA6COyUNJ2dDJAs4lcnOsRn0ueCVS2
# lQhLXkyj8xnOMsrwqqek33vYLPFLsSYhDHR+X75ZU1OGqF25vlvLaJ5BO4v9fa7/
# EtQNlT4lVcTS26cCAwEAAaOCAbswggG3MB8GA1UdIwQYMBaAFFrEuXsqCqOl6nED
# wGD5LfZldQ5YMB0GA1UdDgQWBBT9kTTlmIqXTQkkVW+YyHMkafbVAzAOBgNVHQ8B
# Af8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYv
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmww
# NaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3Mt
# ZzEuY3JsMEIGA1UdIAQ7MDkwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0
# dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwgYQGCCsGAQUFBwEBBHgwdjAkBggr
# BgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJo
# dHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElE
# Q29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOC
# AQEAnJVx0Q4Q8ia/NXog9MfgC9OK8g8HSk/j6KCmIDsXOxIUZ+faIMTTu4wKUD0l
# aNCNdWX5us0OiukT1ocQ3p35X9rzMe2tLSGZaR4athfKa2gDMpCu0xwPRdwFMSG/
# 5VhxmQqHJMg/xkLyXpNy0guTY5MEG57m1mBf12Uw81/5iRlNg3BtThmpnn30dVd8
# Tgx3FBPytz1stW84scG0AkAPehijQlqMOP5hFt9XLqRx0qPl1cZo0wKAvd8pZRPX
# +Cnb+hdvR20ZNBqPGJ/Fo1nixP4bcVv/K7iI9Cp1pBvmMJmQ/p03oGFOkbwFfR2y
# SWFBCH4Vctw3iuTpWPYpmtrHbTCCBTAwggQYoAMCAQICEAQJGBtf1btmdVNDtW+V
# UAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGln
# aUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTEzMTAyMjEyMDAwMFoXDTI4MTAy
# MjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZ
# MBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hB
# MiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wondsy13Hq
# dp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawOeSg6funRZ9PG+yknx9N7I5TkkSOW
# kHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJRdQtoaPpiCwgla4cSocI3wz14k1g
# GL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5sy350OTYN
# kO/ktU6kqepqCquE86xnTrXE94zRICUj6whkPlKWwfIPEvTFjg/BougsUfdzvL2F
# sWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJMBIGA1Ud
# EwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUF
# BwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGln
# aWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2
# hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290
# Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRB
# c3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARIMEYwOAYKYIZIAYb9bAACBDAqMCgG
# CCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAoGCGCGSAGG
# /WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSMEGDAWgBRF
# 66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQsFAAOCAQEAPuwNWiSz8yLR
# FcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1D1G0bqWO
# WuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3nEZOXP+QsRsHDpEV+7qvtVHCjSSu
# JMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4xRJon89A
# y0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOwjNXpDd32
# ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah
# 5xikmmRR7zCCBmowggVSoAMCAQICEAMBmgI6/1ixa9bV6uYX8GYwDQYJKoZIhvcN
# AQEFBQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJl
# ZCBJRCBDQS0xMB4XDTE0MTAyMjAwMDAwMFoXDTI0MTAyMjAwMDAwMFowRzELMAkG
# A1UEBhMCVVMxETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBU
# aW1lc3RhbXAgUmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEAo2Rd/Hyz4II14OD2xirmSXU7zG7gU6mfH2RZ5nxrf2uMnVX4kuOe1VpjWwJJ
# UNmDzm9m7t3LhelfpfnUh3SIRDsZyeX1kZ/GFDmsJOqoSyyRicxeKPRktlC39RKz
# c5YKZ6O+YZ+u8/0SeHUOplsU/UUjjoZEVX0YhgWMVYd5SEb3yg6Np95OX+Koti1Z
# AmGIYXIYaLm4fO7m5zQvMXeBMB+7NgGN7yfj95rwTDFkjePr+hmHqH7P7IwMNlt6
# wXq4eMfJBi5GEMiN6ARg27xzdPpO2P6qQPGyznBGg+naQKFZOtkVCVeZVjCT88lh
# zNAIzGvsYkKRrALA76TwiRGPdwIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeA
# MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAE
# ggG2MIIBsjCCAaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8v
# d3d3LmRpZ2ljZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5
# ACAAdQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABl
# ACAAYwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAg
# AG8AZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABh
# AG4AZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwBy
# AGUAZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBp
# AGwAaQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABl
# AGQAIABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJ
# YIZIAYb9bAMVMB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1Ud
# DgQWBBRhWk0ktkkynUoqeRqDS/QeicHKfTB9BgNVHR8EdjB0MDigNqA0hjJodHRw
# Oi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4
# oDagNIYyaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElE
# Q0EtMS5jcmwwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUA
# A4IBAQCdJX4bM02yJoFcm4bOIyAPgIfliP//sdRqLDHtOhcZcRfNqRu8WhY5AJ3j
# bITkWkD73gYBjDf6m7GdJH7+IKRXrVu3mrBgJuppVyFdNC8fcbCDlBkFazWQEKB7
# l8f2P+fiEUGmvWLZ8Cc9OB0obzpSCfDscGLTYkuw4HOmksDTjjHYL+NtFxMG7uQD
# thSr849Dp3GdId0UyhVdkkHa+Q+B0Zl0DSbEDn8btfWg8cZ3BigV6diT5VUW8LsK
# qxzbXEgnZsijiwoc5ZXarsQuWaBh3drzbaJh6YoLbewSGL33VVRAA5Ira8JRwgpI
# r7DUbuD0FAo6G+OPPcqvao173NhEMIIGzTCCBbWgAwIBAgIQBv35A5YDreoACus/
# J7u6GzANBgkqhkiG9w0BAQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGln
# aUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtE
# aWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMjEx
# MTEwMDAwMDAwWjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBB
# c3N1cmVkIElEIENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDo
# gi2Z+crCQpWlgHNAcNKeVlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561JBXCmLm0
# d0ncicQK2q/LXmvtrbBxMevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFIgHweog+S
# DlDJxofrNj/YMMP/pvf7os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C6Ds97bFB
# o+0/vtuVSMTuHrPyvAwrmdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4Grq2Eb0iE
# m09AufFM8q+Y+/bOQF1c9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrhofrfVdwon
# VnwPYqQ/MhRglf0HBKIJAgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMCAYYwOwYD
# VR0lBDQwMgYIKwYBBQUHAwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYBBQUHAwQG
# CCsGAQUFBwMIMIIB0gYDVR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEEMIIBpDA6
# BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBv
# c2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAg
# AG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBz
# AHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABo
# AGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABo
# AGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBu
# AHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAg
# AGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQBy
# AGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUw
# EgYDVR0TAQH/BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCB
# gQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lD
# ZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNl
# cnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUFQAS
# KxOYspkH7R7for5XDStnAs0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNt
# yA8wDQYJKoZIhvcNAQEFBQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFHq1Zce9UN
# C0Gz7+x1H3Q48rJcYaKclcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR67s2hHfM
# JKXzBBlVqefj56tizfuLLZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPkmzmnxPXO
# HXh2lCVz5Cqrz5x2S+1fwksW5EtwTACJHvzFebxMElf+X+EevAJdqP77BzhPDcZd
# kbkPZ0XN1oPt55INjbFpjE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718CqDuLAhV
# hSK46xgaTfwqIa1JMYNHlXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQ7MIIENwIB
# ATCBhjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBAhADViTO4HBjoJNSwH9//cwJMAkGBSsO
# AwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBSlr3SuBVazSPWdcs5qv703+r5F2jANBgkqhkiG9w0BAQEFAASC
# AQBAePLKC6oYlNN82o3JW55kHkDQGsU7GF0rPUfgC4f1kQShVVPGcIk92WkK7ONh
# +F9PZwtM8jd5W0cI0sTyC+54thHvz4MtOgfL2hB7O/bvWF6VhXL+0JmTVcU2vK8A
# Z8sHzK0CL4qn1MXao8VVhvKMFZ2em5p/9UxqKXgpjSYqKkZ0ONLJ9okBqGd8ONdt
# 4CQE2TfxzMR5vPig/xI7zwNVJ99rJlxf58UFR7VU3cSuxH4HKJuGHqn9uZV5BWkr
# tt4LnEwKFl5eFcRdZO+30sSDrShnkMopjwh9KIMG2J1ItFJVeg7Gv4v/GLtSLCuf
# CoyR0JHa9Ro5kM+grl3trT6GoYICDzCCAgsGCSqGSIb3DQEJBjGCAfwwggH4AgEB
# MHYwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJ
# RCBDQS0xAhADAZoCOv9YsWvW1ermF/BmMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0B
# CQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNTA3MDEwOTQ3MTBaMCMG
# CSqGSIb3DQEJBDEWBBRMYy4csIbSi3wNpVvDHn+MZ0H6ezANBgkqhkiG9w0BAQEF
# AASCAQBPn960B8dluVJeMb92RJcEZFNEnJjoyh+h4quncgSVLK4kXio3CXexF0KE
# 9UgkE07wb1RLJJqROMtsEaGsVsINk34TLd1sbJHDpZISt2MdDVSXj/J6GDpvwkjq
# RZnT44lQhCtnzAyUfvLiqN05sFT1DWyVEv4RDLUmAfYhNoLpHutqolLvQMS4BKRi
# b0PReff/F0PxF/BanKEhWmpF9nmhlYCIEs8+OLdliAChjoEP5mW7eFk9q0q46E9r
# MM0dsVxQu8SRUQvOeZuE/hAtDLOOqA4SQ/EDwCf08JeEpWfH0ZvyA9cVeBDSJtsl
# 3tLSZ4tW5aYPnQF6GQoLcqls5UjA
# SIG # End signature block
