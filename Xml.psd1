@{

# Script module or binary module file associated with this manifest.
RootModule = 'Xml.psm1'

# Version number of this module.
ModuleVersion = '7.1'

# ID used to uniquely identify this module
GUID = '9f58549b-020b-4908-b8ce-8b37ed8f5a2b'

# Author of this module
Author = 'Joel "Jaykul" Bennett'

# Company or vendor of this module
CompanyName = 'HuddledMasses.org'

# Copyright statement for this module
Copyright = '(c) 2018 Joel Bennett. All rights reserved.'

# Description of the functionality provided by this module
Description = 'A module providing converters for HTML to XML, various core XML commands and a DSL for generating XML documents.'

# Minimum version of the Windows PowerShell engine required by this module
# PowerShellVersion = ''

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module
FunctionsToExport = 'New-XDocument', 'New-XAttribute', 'New-XElement', 'Remove-XmlNamespace', 'Remove-XmlElement', 'Get-XmlContent', 'Set-XmlContent', 'Convert-Xml', 'Select-Xml', 'Update-Xml', 'Format-Xml', 'ConvertTo-CliXml', 'ConvertFrom-CliXml', 'Import-Html', 'ConvertFrom-Html', 'ConvertFrom-XmlDsl'

# Cmdlets to export from this module
# CmdletsToExport =

# Variables to export from this module
# VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = 'cvxml','epx','epxml','Export-Xml','fx','fxml','Get-Xml','gx','gxml','Import-Xml','ipx','ipxml','New-Xml','New-XmlAttribute','New-XmlElement','rmns','rmxns','Set-Xml','slx','slxml','sx','sxml','ux','uxml','xa','xe','xml'

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = @("Xml")

        # A URL to the license for this module.
        LicenseUri = 'https://raw.githubusercontent.com/Jaykul/Xml/master/LICENSE'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/Jaykul/Xml'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        ReleaseNotes = 'Fix Add-XNamespace for documents requiring no namespace at all.'

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}
