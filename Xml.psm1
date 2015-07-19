#requires -version 2.0

# Improves over the built-in Select-XML by leveraging Remove-XmlNamespace http`://poshcode.org/1492 
# to provide a -RemoveNamespace parameter -- if it's supplied, all of the namespace declarations 
# and prefixes are removed from all XML nodes (by an XSL transform) before searching. 
# IMPORTANT: returned results *will not* have namespaces in them, even if the input XML did. 

# Also, only raw XmlNodes are returned from this function, so the output isn't completely compatible 
# with the built in Select-Xml. It's equivalent to using Select-Xml ... | Select-Object -Expand Node

# Version History:
# Select-Xml 2.0 This was the first script version I wrote.
#                it didn't function identically to the built-in Select-Xml with regards to parameter parsing
# Select-Xml 2.1 Matched the built-in Select-Xml parameter sets, it's now a drop-in replacement 
#                BUT only if you were using the original with: Select-Xml ... | Select-Object -Expand Node
# Select-Xml 2.2 Fixes a bug in the -Content parameterset where -RemoveNamespace was *presumed*
# Version    3.0 Added New-XDocument and associated generation functions for my XML DSL
# Version    3.1 Fixed a really ugly bug in New-XDocument in 3.0 which I should not have released
# Version    4.0 Never content to leave well enough alone, I've completely reworked New-XDocument
# Version    4.1 Tweaked namespaces again so they don't cascade down when they shouldn't. Got rid of the unnecessary stack.
# Version    4.2 Tightened xml: only cmdlet, function, and external scripts, with "-" in their names are exempted from being converted into xml tags.
#                Fixed some alias error messages caused when PSCX is already loaded (we overwrite their aliases for cvxml and fxml)
# Version    4.3 Added a Path parameter set to Format-Xml so you can specify xml files for prety printing
# Version    4.5 Fixed possible [Array]::Reverse call on a non-array in New-XElement (used by New-XDocument)
#                Work around possible variable slipping on null values by:
#                1) allowing -param:$value syntax (which doesn't fail when $value is null)
#                2) testing for -name syntax on the value and using it as an attribute instead
# Version    4.6 Added -Arguments to Convert-Xml so that you can pass arguments to XSLT transforms!
#                Note: when using strings for xslt, make sure you single quote them or escape the $ signs.
# Version    4.7 Fixed a typo in the namespace parameter of Select-Xml
# Version    4.8 Fixed up some uses of Select-Xml -RemoveNamespace
# Version    5.0 Added Update-Xml to allow setting xml attributes or node content
# Version    6.0 Major cleanup, breaking changes.
#       - Added Get-XmlContent and Set-XmlContent for loading/saving XML from files or strings
#       - Removed Path and Content parameters from the other functions (it greatly simplifies thost functions, and makes the whole thing more maintainable)
#       - Updated Update-Xml to support adding nodes "before" and "after" other nodes, and to support "remove"ing nodes
# Version    6.1 Update for PowerShell 3.0

function Add-Accelerator {
<#
   .Synopsis
      Add a type accelerator to the current session
   .Description
      The Add-Accelerator function allows you to add a simple type accelerator (like [regex]) for a longer type (like [System.Text.RegularExpressions.Regex]).
   .Example
      Add-Accelerator list System.Collections.Generic.List``1
      $list = New-Object list[string]
      
      Creates an accelerator for the generic List[T] collection type, and then creates a list of strings.
   .Example
      Add-Accelerator "List T", "GList" System.Collections.Generic.List``1
      $list = New-Object "list t[string]"
      
      Creates two accelerators for the Generic List[T] collection type.
   .Parameter Accelerator
      The short form accelerator should be just the name you want to use (without square brackets).
   .Parameter Type
      The type you want the accelerator to accelerate (without square brackets)
   .Notes
      When specifying multiple values for a parameter, use commas to separate the values. 
      For example, "-Accelerator string, regex".
      
      PowerShell requires arguments that are "types" to NOT have the square bracket type notation, because of the way the parsing engine works.  You can either just type in the type as System.Int64, or you can put parentheses around it to help the parser out: ([System.Int64])

      Also see the help for Get-Accelerator and Remove-Accelerator
   .Link
      http://huddledmasses.org/powershell-2-ctp3-custom-accelerators-finally/
#>
[CmdletBinding()]
param(
   [Parameter(Position=0,ValueFromPipelineByPropertyName=$true)]
   [Alias("Key","Name")]
   [string[]]$Accelerator
,
   [Parameter(Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
   [Alias("Value","FullName")]
   [type]$Type
)
process {
   # add a user-defined accelerator  
   foreach($a in $Accelerator) { 
      if($xlr8r::AddReplace) { 
         $xlr8r::AddReplace( $a, $Type) 
      } else {
         $null = $xlr8r::Remove( $a )
         $xlr8r::Add( $a, $Type)
      }
      trap [System.Management.Automation.MethodInvocationException] {
         if($xlr8r::get.keys -contains $a) {
            if($xlr8r::get[$a] -ne $Type) {
               Write-Error "Cannot add accelerator [$a] for [$($Type.FullName)]`n                  [$a] is already defined as [$($xlr8r::get[$a].FullName)]"
            }
            Continue;
         } 
         throw
      }
   }
}
}

&{ 
$local:xlr8r = [psobject].assembly.gettype("System.Management.Automation.TypeAccelerators")
$local:xlinq = [Reflection.Assembly]::Load("System.Xml.Linq, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
$xlinq.GetTypes() | ? { $_.IsPublic -and !$_.IsSerializable -and $_.Name -ne "Extensions" -and !$xlr8r::Get[$_.Name] } | Add-Accelerator

Add-Accelerator "Dictionary" "System.Collections.Generic.Dictionary``2, mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Add-Accelerator "Dictionary", "System.Collections.Generic.Dictionary``2, mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Add-Accelerator "PSParser", "System.Management.Automation.PSParser, System.Management.Automation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
}


function Get-XmlContent {
#.Synopsis
#   Load an XML file as an XmlDocument
param(
    # Specifies a string that contains the XML to load, or a path to a file which has the XML to load (wildcards are permitted).
    [Parameter(Position=1,Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [Alias("PSPath","Path")]
    [String[]]$Content
,
    # If set, loads XML with all namespace qualifiers removed, and all entities expanded.
    [Alias("Rn","Rm")]
    [Switch]$RemoveNamespace
)
begin {
    [Text.StringBuilder]$XmlContent = [String]::Empty
    [bool]$Path = $true
}
process {
    if($Path -and ($Path = Test-Path @($Content)[0] -EA 0)) { 
        foreach($file in Resolve-Path $Content) {
            $xml = New-Object System.Xml.XmlDocument;
            if($file.Provider.Name -eq "FileSystem") {
                Write-Verbose $file.ProviderPath
                $xml.Load( $file.ProviderPath )
            } else {
                $ofs = "`n"
                $xml.LoadXml( ([String](Get-Content $file)) )
            }
            if($RemoveNamespace) {
                [System.Xml.XmlNode[]]$Xml = @(Remove-XmlNamespace -Xml $node)
            }
            Write-Output $xml
        }
    } else {
        # If the "path" parameter isn't actually a path, assume that it's actually content
        foreach($line in $content) {
            $null = $XmlContent.AppendLine( $line )
        }
    }
}
end {
    if(!$Path) {
        $xml = New-Object System.Xml.XmlDocument; 
        $xml.LoadXml( $XmlContent.ToString() )
        if($RemoveNamespace) {
            $Xml = @(Remove-XmlNamespace -Xml $xml)
        }
        Write-Output $xml
    }
}}


Set-Alias Import-Xml Get-XmlContent
Set-Alias ipxml Get-XmlContent
Set-Alias ipx Get-XmlContent
Set-Alias Get-Xml Get-XmlContent
Set-Alias gxml Get-XmlContent
Set-Alias gx Get-XmlContent

function Set-XmlContent {
param(
    [Parameter(Mandatory=$true, Position=1)]
    [Alias("PSPath")]
    [String]$Path
,
    # Specifies one or more XML nodes to search.
    [Parameter(Position=5,ParameterSetName="Xml",Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [Alias("Node")]
    [Xml]$Xml
)
process {
    $xml.Save( $Path )
}
}

Set-Alias Export-Xml Set-XmlContent
Set-Alias epxml Set-XmlContent
Set-Alias epx Set-XmlContent
Set-Alias Set-Xml Set-XmlContent
Set-Alias sxml Set-XmlContent
Set-Alias sx Set-XmlContent

function Format-Xml {
#.Synopsis
#   Pretty-print formatted XML source
#.Description
#   Runs an XmlDocument through an auto-indenting XmlWriter
#.Parameter Xml
#   The Xml Document
#.Parameter Path
#   The path to an xml document (on disc or any other content provider).
#.Parameter Indent
#   The indent level (defaults to 2 spaces)
#.Example
#   [xml]$xml = get-content Data.xml
#   C:\PS>Format-Xml $xml
#.Example
#   get-content Data.xml | Format-Xml
#.Example
#   Format-Xml C:\PS\Data.xml
#.Example
#   ls *.xml | Format-Xml
#
[CmdletBinding()]
param(
   [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="Document")]
   [xml]$Xml
,
   [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName="File")]
   [Alias("PsPath")]
   [string]$Path
,
   [Parameter(Mandatory=$false)]
   $Indent=2
)
process {
   ## Load from file, if necessary
   if($Path) { [xml]$xml = Get-Content $Path }
   
   $StringWriter = New-Object System.IO.StringWriter
   $XmlWriter = New-Object System.Xml.XmlTextWriter $StringWriter
   $xmlWriter.Formatting = "indented"
   $xmlWriter.Indentation = $Indent
   $xml.WriteContentTo($XmlWriter)
   $XmlWriter.Flush()
   $StringWriter.Flush()
   Write-Output $StringWriter.ToString()
}}
Set-Alias fxml Format-Xml -EA 0
Set-Alias fx   Format-Xml -EA 0

function Select-XmlNodeInternal {
[CmdletBinding()]
param([Xml.XmlNode[]]$Xml, [String[]]$XPath, [Hashtable]$NamespaceManager)
begin {
    Write-Verbose "XPath = $($XPath -join ',')"
    foreach($node in $xml) {
        if($NamespaceManager) {
            $nsManager = new-object System.Xml.XmlNamespaceManager $node.NameTable
            foreach($ns in $NamespaceManager.GetEnumerator()) {
                $nsManager.AddNamespace( $ns.Key, $ns.Value )
            }
            Write-Verbose "Names = $($nsManager | % { @{ $_ = $nsManager.LookupNamespace($_) } } | Out-String)"
        }
        foreach($path in $xpath) {
            $node.SelectNodes($path, $nsManager)
        }
    }
}}

function Select-Xml {
#.Synopsis
#  The Select-XML cmdlet lets you use XPath queries to search for text in XML strings and documents. Enter an XPath query, and use the Content, Path, or Xml parameter to specify the XML to be searched.
#.Description
#  Improves over the built-in Select-XML by leveraging Remove-XmlNamespace to provide a -RemoveNamespace parameter -- if it's supplied, all of the namespace declarations and prefixes are removed from all XML nodes (by an XSL transform) before searching.  
#  
#  However, only raw XmlNodes are returned from this function, so the output isn't currently compatible with the built in Select-Xml, but is equivalent to using Select-Xml ... | Select-Object -Expand Node
#
#  Also note that if the -RemoveNamespace switch is supplied the returned results *will not* have namespaces in them, even if the input XML did, and entities get expanded automatically.
[CmdletBinding(DefaultParameterSetName="Xml")]
param(
    # Specifies an XPath search query. The query language is case-sensitive. This parameter is required.
    [Parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false)]
    [ValidateNotNullOrEmpty()]
    [Alias("Query")]
    [String[]]$XPath
,
    # Specifies a string that contains the XML to search. You can also pipe strings to Select-XML.
    [Parameter(ParameterSetName="Content",Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]]$Content
,
    # Specifies the path and file names of the XML files to search.  Wildcards are permitted.
    [Parameter(Position=5,ParameterSetName="Path",Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [Alias("PSPath")]
    [String[]]$Path
,
    # Specifies one or more XML nodes to search.
    [Parameter(Position=5,ParameterSetName="Xml",Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [Alias("Node")]
    [System.Xml.XmlNode[]]$Xml
,
    # Specifies a hash table of the namespaces used in the XML. Use the format @{<namespaceName> = <namespaceUri>}.
    [Parameter(Position=10,Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [Alias("Ns")]
    [Hashtable]$Namespace
,
    # Allows the execution of XPath queries without namespace qualifiers. 
    # 
    # If you specify the -RemoveNamespace switch, all namespace declarations and prefixes are actually removed from the Xml before the XPath search query is evaluated, and your XPath query should therefore NOT contain any namespace prefixes.
    # 
    # Note that this means that the returned results *will not* have namespaces in them, even if the input XML did, and entities get expanded automatically.
    [Alias("Rn","Rm")]
    [Switch]$RemoveNamespace
)
begin {
    $NSM = $Null; if($PSBoundParameters.ContainsKey("Namespace")) { $NSM = $Namespace }
    $XmlNodes = New-Object System.Xml.XmlNode[] 1
    if($PSCmdlet.ParameterSetName -eq "Content") {
        $XmlNodes = ConvertTo-Xml $Content -RemoveNamespace:$RemoveNamespace
        Select-XmlNodeInternal $XmlNodes $XPath $NSM
    }
}
process {
    switch($PSCmdlet.ParameterSetName) {
        "Path" {
            $node = ConvertTo-Xml $Path -RemoveNamespace:$RemoveNamespace
            Select-XmlNodeInternal $node $XPath $NSM
        }
        "Xml" {
            foreach($node in $Xml) {
                if($RemoveNamespace) {
                   [Xml]$node = Remove-XmlNamespace -Xml $node
                }
                Select-XmlNodeInternal $node $XPath $NSM
            }
        }
    }
}}
Set-Alias slxml Select-Xml -EA 0
Set-Alias slx Select-Xml -EA 0


function Update-Xml {
#.Synopsis
#  The Update-XML cmdlet lets you use XPath queries to replace text in nodes in XML documents. Enter an XPath query, and use the Content, Path, or Xml parameter to specify the XML to be searched.
#.Description
#  Allows you to update an attribute value, xml node contents, etc.
#
#.Notes
#  We still need to implement RemoveNode and RemoveAttribute and even ReplaceNode
[CmdletBinding(DefaultParameterSetName="Set")]
param(
    # Specifies an XPath for an element where you want to insert the new node.
    [Parameter(ParameterSetName="Before",Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [Switch]$Before
,
    # Specifies an XPath for an element where you want to insert the new node.
    [Parameter(ParameterSetName="After",Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [Switch]$After
,
    # If set, the new value will be added as a new child of the node identified by the XPath
    [Parameter(ParameterSetName="Append",Mandatory=$true)]
    [Switch]$Append
,
    # If set, the node identified by the XPath will be removed instead of set
    [Parameter(ParameterSetName="Remove",Mandatory=$true)]
    [Switch]$Remove
,
    # If set, the node identified by the XPath will be Replace instead of set
    [Parameter(ParameterSetName="Replace",Mandatory=$true)]
    [Switch]$Replace
,
    # Specifies an XPath for the node to update. This could be an element node *or* an attribute node (remember: //element/@attribute )
    [Parameter(Position=1,Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]]$XPath
,
    # The new value to place in the xml
    [Parameter(Position=2,Mandatory=$true,ValueFromPipeline=$false)]
    [ValidateNotNullOrEmpty()]
    [String]$Value
,
    # Specifies one or more XML nodes to search.
    [Parameter(Position=5,Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [Alias("Node")]
    [System.Xml.XmlNode[]]$Xml
,   
    # Specifies a hash table of the namespaces used in the XML. Use the format @{<namespaceName> = <namespaceUri>}.
    [Parameter(Position=10,Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [Alias("Ns")]
    [Hashtable]$Namespace
,   
    # Output the XML documents after adding updating them
    [Switch]$Passthru
)
process
{
    foreach($XmlNode in $Xml) {
        $select = @{}
        $select.Xml = $XmlNode
        $select.XPath = $XPath
        if($Namespace) {  
            $select.Namespace = $Namespace
        }
        $document =
            if($XmlNode -is [System.Xml.XmlDocument]) {
                $XmlNode
            } else { 
                $XmlNode.get_OwnerDocument()
            }
        if($xValue = $Value -as [Xml]) {
            $xValue = $document.ImportNode($xValue.SelectSingleNode("/*"), $true)
        }
        $nodes = Select-Xml @Select | Where-Object { $_ }

        if(@($nodes).Count -eq 0) { Write-Warning "No nodes matched your XPath, nothing will be updated" }
        
        foreach($node in $nodes) {
            $select.XPath = "$XPath/parent::*"
            $parent = Select-Xml @Select
            if(!$xValue) {
                if($node -is [System.Xml.XmlAttribute] -and $Value.Contains("=")) {
                    $aName, $aValue = $Value.Split("=",2)
                    if($aName.Contains(":")){
                        $ns,$name = $aName.Split(":",2)
                        $xValue = $document.CreateAttribute( $name, $Namespace[$ns] )
                    } else {
                        $xValue = $document.CreateAttribute( $aName )
                    }
                    $xValue.Value = $aValue
                }
            }
            
            switch($PSCmdlet.ParameterSetName) {
                "Before" {
                    $null = $parent.InsertBefore( $xValue, $node )
                }
                "After" {
                    $null = $parent.InsertAfter( $xValue, $node )
                }
                "Append" {
                    $null = $parent.AppendChild( $xValue )
                }
                "Remove" {
                    $null = $parent.RemoveChild( $node )
                }
                "Replace" {
                    if(!$xValue) {
                        $xValue = $document.CreateTextNode( $Value )
                    }
                    $null = $parent.ReplaceChild( $xValue, $node )
                }
                "Set" {
                    if(!$xValue -and $node."#text") {
                        $node."#text" = $Value
                    } else {
                        if($node -is [System.Xml.XmlElement]) {
                            if(!$xValue) {
                                $xValue = $document.CreateTextNode( $Value )
                            }
                            $null = $node.set_innerXml("")
                            $null = $node.AppendChild($xValue)
                        }
                        elseif($node -is [System.Xml.XmlAttribute]) {
                            $node.Value = $Value
                        } else {
                            Write-Warning "$XPath selects a node of type $($node.GetType()), which we haven't handled. Please add that handler!"
                        }
                    }
                }
            }
        }
        if($Passthru) {
            Write-Output $XmlNode
        }
    }
}}
Set-Alias uxml Update-Xml -EA 0
Set-Alias ux Update-Xml -EA 0

function Convert-Node {
#.Synopsis 
# Convert a single XML Node via XSL stylesheets
[CmdletBinding(DefaultParameterSetName="Reader")]
param(
   [Parameter(ParameterSetName="ByNode",Mandatory=$true,ValueFromPipeline=$true)]
   [System.Xml.XmlNode]$Node
,
   [Parameter(ParameterSetName="Reader",Mandatory=$true,ValueFromPipeline=$true)]
   [System.Xml.XmlReader]$XmlReader
,
   [Parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false)]
   [System.Xml.Xsl.XslCompiledTransform]$StyleSheet
,
   [Parameter(Position=2,Mandatory=$false)]
   [Alias("Parameters")]
   [hashtable]$Arguments
)
PROCESS {
   if($PSCmdlet.ParameterSetName -eq "ByNode") {
      $XmlReader = New-Object Xml.XmlNodeReader $node
   }

   $output = New-Object IO.StringWriter
   $argList = $null
   
   if($Arguments) {
      $argList = New-Object System.Xml.Xsl.XsltArgumentList
      foreach($arg in $Arguments.GetEnumerator()) {
         $namespace, $name = $arg.Key -split ":"
         ## Fix namespace
         if(!$name) { 
            $name = $Namespace
            $namespace = ""
         }
         
         Write-Verbose "ns:$namespace name:$name value:$($arg.Value)"
         $argList.AddParam($name,"$namespace",$arg.Value)
      }
   }
   
   $StyleSheet.Transform( $XmlReader, $argList, $output )
   Write-Output $output.ToString()
}
}

function Convert-Xml {
#.Synopsis
#   The Convert-XML function lets you use Xslt to transform XML strings and documents.
#.Description
#   Documentation TODO
[CmdletBinding(DefaultParameterSetName="Xml")]
param(
    # Specifies one or more XML nodes to process.
    [Parameter(Position=1,ParameterSetName="Xml",Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [Alias("Node")]
    [System.Xml.XmlNode[]]$Xml
,   
    # Specifies an Xml StyleSheet to transform with...
    [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false)]
    [ValidateNotNullOrEmpty()]
    [Alias("StyleSheet")]
    [String]$Xslt
,
    # Specify arguments to the XSL Transformation
    [Alias("Parameters")]
    [hashtable]$Arguments
)
begin { 
   $StyleSheet = New-Object System.Xml.Xsl.XslCompiledTransform
   if(Test-Path $Xslt -EA 0) { 
      Write-Verbose "Loading Stylesheet from $(Resolve-Path $Xslt)"
      $StyleSheet.Load( (Resolve-Path $Xslt) )
   } else {
      $OFS = "`n"
      Write-Verbose "$Xslt"
      $StyleSheet.Load(([System.Xml.XmlReader]::Create((New-Object System.IO.StringReader $Xslt))))
   }
}
process {
   foreach($node in $Xml) {
      Convert-Node -Xml (New-Object Xml.XmlNodeReader $node) $StyleSheet $Arguments
   }
}
}
Set-Alias cvxml Convert-Xml -EA 0

function Remove-XmlNamespace {
#.Synopsis
#  Removes namespace definitions and prefixes from xml documents
#.Description
#  Runs an xml document through an XSL Transformation to remove namespaces from it if they exist.
#  Entities are also naturally expanded
#.Parameter Content
#  Specifies a string that contains the XML to transform.
#.Parameter Path
#  Specifies the path and file names of the XML files to transform. Wildcards are permitted.
#
#  There will bne one output document for each matching input file.
#.Parameter Xml
#  Specifies one or more XML documents to transform
[CmdletBinding(DefaultParameterSetName="Xml")]
PARAM(
   [Parameter(Position=1,ParameterSetName="Xml",Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
   [ValidateNotNullOrEmpty()]
   [Alias("Node")]
   [System.Xml.XmlNode[]]$Xml
)
BEGIN { 
   $StyleSheet = New-Object System.Xml.Xsl.XslCompiledTransform
   $StyleSheet.Load(([System.Xml.XmlReader]::Create((New-Object System.IO.StringReader @"
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   <xsl:output method="xml" indent="yes"/>
   <xsl:template match="/|comment()|processing-instruction()">
      <xsl:copy>
         <xsl:apply-templates/>
      </xsl:copy>
   </xsl:template>

   <xsl:template match="*">
      <xsl:element name="{local-name()}">
         <xsl:apply-templates select="@*|node()"/>
      </xsl:element>
   </xsl:template>

   <xsl:template match="@*">
      <xsl:attribute name="{local-name()}">
         <xsl:value-of select="."/>
      </xsl:attribute>
   </xsl:template>
</xsl:stylesheet>
"@))))
   [Text.StringBuilder]$XmlContent = [String]::Empty 
}
PROCESS {
   $Xml | Convert-Node $StyleSheet
}
}
Set-Alias rmns Remove-XmlNamespace -EA 0
Set-Alias rmxns Remove-XmlNamespace -EA 0

######## Helper functions for working with CliXml

function ConvertFrom-CliXml {
   param(
      [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
      [ValidateNotNullOrEmpty()]
      [String[]]$InputObject
   )
   begin
   {
      $OFS = "`n"
      [String]$xmlString = ""
   }
   process
   {
      $xmlString += $InputObject
   }
   end
   {
      $type = [psobject].assembly.gettype("System.Management.Automation.Deserializer")
      $ctor = $type.getconstructor("instance,nonpublic", $null, @([xml.xmlreader]), $null)
      $sr = new-object System.IO.StringReader $xmlString
      $xr = new-object System.Xml.XmlTextReader $sr
      $deserializer = $ctor.invoke($xr)
      $method = @($type.getmethods("nonpublic,instance") | where-object {$_.name -like "Deserialize"})[1]
      $done = $type.getmethod("Done", [System.Reflection.BindingFlags]"nonpublic,instance")
      while (!$done.invoke($deserializer, @()))
      {
         try {
            $method.invoke($deserializer, "")
         } catch {
            write-warning "Could not deserialize ${string}: $_"
         }
      }
      $xr.Close()
      $sr.Dispose()
   }
}

function ConvertTo-CliXml {
   param(
      [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
      [ValidateNotNullOrEmpty()]
      [PSObject[]]$InputObject
   )
   begin {
      $type = [psobject].assembly.gettype("System.Management.Automation.Serializer")
      $ctor = $type.getconstructor("instance,nonpublic", $null, @([System.Xml.XmlWriter]), $null)
      $sw = new-object System.IO.StringWriter
      $xw = new-object System.Xml.XmlTextWriter $sw
      $serializer = $ctor.invoke($xw)
      $method = $type.getmethod("Serialize", "nonpublic,instance", $null, [type[]]@([object]), $null)
      $done = $type.getmethod("Done", [System.Reflection.BindingFlags]"nonpublic,instance")
   }
   process {
      try {
         [void]$method.invoke($serializer, $InputObject)
      } catch {
         write-warning "Could not serialize $($InputObject.gettype()): $_"
      }
   }
   end {    
      [void]$done.invoke($serializer, @())
      $sw.ToString()
      $xw.Close()
      $sw.Dispose()
   }
}


######## From here down is all the code related to my XML DSL:

function New-XDocument {
#.Synopsis
#   Creates a new XDocument (the new xml document type)
#.Description
#  This is the root for a new XML mini-dsl, akin to New-BootsWindow for XAML
#  It creates a new XDocument, and takes scritpblock(s) to define it's contents
#.Parameter root
#   The root node name
#.Parameter version
#   Optional: the XML version. Defaults to 1.0
#.Parameter encoding
#   Optional: the Encoding. Defaults to UTF-8
#.Parameter standalone
#  Optional: whether to specify standalone in the xml declaration. Defaults to "yes"
#.Parameter args
#   this is where all the dsl magic happens. Please see the Examples. :)
#
#.Example
# [string]$xml = New-XDocument rss -version "2.0" {
#    channel {
#       title {"Test RSS Feed"}
#       link {"http`://HuddledMasses.org"}
#       description {"An RSS Feed generated simply to demonstrate my XML DSL"}
#       item {
#          title {"The First Item"}
#          link {"http`://huddledmasses.org/new-site-new-layout-lost-posts/"}
#          guid -isPermaLink true {"http`://huddledmasses.org/new-site-new-layout-lost-posts/"}
#          description {"Ema Lazarus' Poem"}
#          pubDate {(Get-Date 10/31/2003 -f u) -replace " ","T"}
#       }
#    }
# }
#
# C:\PS>$xml.Declaration.ToString()  ## I can't find a way to have this included in the $xml.ToString()
# C:\PS>$xml.ToString()
#
# <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
# <rss version="2.0">
#   <channel>
#     <title>Test RSS Feed</title>
#     <link>http ://HuddledMasses.org</link>
#     <description>An RSS Feed generated simply to demonstrate my XML DSL</description>
#     <item>
#       <title>The First Item</title>
#       <link>http ://huddledmasses.org/new-site-new-layout-lost-posts/</link>
#       <guid isPermaLink="true">http ://huddledmasses.org/new-site-new-layout-lost-posts/</guid>
#       <description>Ema Lazarus' Poem</description>
#       <pubDate>2003-10-31T00:00:00Z</pubDate>
#     </item>
#   </channel>
# </rss>
#
#
# Description
# -----------
# This example shows the creation of a complete RSS feed with a single item in it. 
#
# NOTE that the backtick in the http`: in the URLs in the input is unecessary, and I added the space after the http: in the URLs  in the output -- these are accomodations to PoshCode's spam filter. Backticks are not need in the input, and spaces do not appear in the actual output.
#
#
#.Example 
# [XNamespace]$atom="http`://www.w3.org/2005/Atom"
# C:\PS>[XNamespace]$dc = "http`://purl.org/dc/elements/1.1"
# 
# C:\PS>New-XDocument ($atom + "feed") -Encoding "UTF-16" -$([XNamespace]::Xml +'lang') "en-US" -dc $dc {
#    title {"Test First Entry"}
#    link {"http`://HuddledMasses.org"}
#    updated {(Get-Date -f u) -replace " ","T"}
#    author {
#       name {"Joel Bennett"}
#       uri {"http`://HuddledMasses.org"}
#    }
#    id {"http`://huddledmasses.org/" }
#
#    entry {
#       title {"Test First Entry"}
#       link {"http`://HuddledMasses.org/new-site-new-layout-lost-posts/" }
#       id {"http`://huddledmasses.org/new-site-new-layout-lost-posts/" }
#       updated {(Get-Date 10/31/2003 -f u) -replace " ","T"}
#       summary {"Ema Lazarus' Poem"}
#       link -rel license -href "http`://creativecommons.org/licenses/by/3.0/" -title "CC By-Attribution"
#       dc:rights { "Copyright 2009, Some rights reserved (licensed under the Creative Commons Attribution 3.0 Unported license)" }
#       category -scheme "http`://huddledmasses.org/tag/" -term "huddled-masses"
#    }
# } | % { $_.Declaration.ToString(); $_.ToString() }
#
# <?xml version="1.0" encoding="UTF-16" standalone="yes"?>
# <feed xml:lang="en-US" xmlns="http ://www.w3.org/2005/Atom">
#   <title>Test First Entry</title>
#   <link>http ://HuddledMasses.org</link>
#   <updated>2009-07-29T17:25:49Z</updated>
#   <author>
#      <name>Joel Bennett</name>
#      <uri>http ://HuddledMasses.org</uri>
#   </author>
#   <id>http ://huddledmasses.org/</id>
#   <entry>
#     <title>Test First Entry</title>
#     <link>http ://HuddledMasses.org/new-site-new-layout-lost-posts/</link>
#     <id>http ://huddledmasses.org/new-site-new-layout-lost-posts/</id>
#     <updated>2003-10-31T00:00:00Z</updated>
#     <summary>Ema Lazarus' Poem</summary>
#     <link rel="license" href="http ://creativecommons.org/licenses/by/3.0/" title="CC By-Attribution" />
#     <dc:rights>Copyright 2009, Some rights reserved (licensed under the Creative Commons Attribution 3.0 Unported license)</dc:rights>
#     <category scheme="http ://huddledmasses.org/tag/" term="huddled-masses" />
#   </entry>
# </feed>
#
#
# Description
# -----------
# This example shows the use of a default namespace, as well as additional specific namespaces for the "dc" namespace. It also demonstrates how you can get the <?xml?> declaration which does not appear in a simple .ToString().
#
# NOTE that the backtick in the http`: in the URLs in the input is unecessary, and I added the space after the http: in the URLs  in the output -- these are accomodations to PoshCode's spam filter. Backticks are not need in the input, and spaces do not appear in the actual output.#
# 
[CmdletBinding()]
Param(
   [Parameter(Mandatory = $true, Position = 0)]
   [System.Xml.Linq.XName]$root
,
   [Parameter(Mandatory = $false)]
   [string]$Version = "1.0"
,
   [Parameter(Mandatory = $false)]
   [string]$Encoding = "UTF-8"
,
   [Parameter(Mandatory = $false)]
   [string]$Standalone = "yes"
,
   [AllowNull()][AllowEmptyString()][AllowEmptyCollection()]
   [Parameter(Position=99, Mandatory = $false, ValueFromRemainingArguments=$true)]
   [PSObject[]]$args
)
BEGIN {
   $script:NameSpaceHash = New-Object 'Dictionary[String,XNamespace]'
   if($root.NamespaceName) {
      $script:NameSpaceHash.Add("", $root.Namespace)
   }
}
PROCESS {
   New-Object XDocument (New-Object XDeclaration $Version, $Encoding, $standalone),(
      New-Object XElement $(
         $root
         while($args) {
            $attrib, $value, $args = $args
            if($attrib -is [ScriptBlock]) {
               # Write-Verbose "Preparsed DSL: $attrib"
               $attrib = ConvertFrom-XmlDsl $attrib
               Write-Verbose "Reparsed DSL: $attrib"
               &$attrib
            } elseif ( $value -is [ScriptBlock] -and "-CONTENT".StartsWith($attrib.TrimEnd(':').ToUpper())) {
               $value = ConvertFrom-XmlDsl $value
               &$value
            } elseif ( $value -is [XNamespace]) {
               New-Object XAttribute ([XNamespace]::Xmlns + $attrib.TrimStart("-").TrimEnd(':')), $value
               $script:NameSpaceHash.Add($attrib.TrimStart("-").TrimEnd(':'), $value)
            } else {
               Write-Verbose "XAttribute $attrib = $value"
               New-Object XAttribute $attrib.TrimStart("-").TrimEnd(':'), $value
            }
         }
      ))
}
}

Set-Alias xml New-XDocument -EA 0
Set-Alias New-Xml New-XDocument -EA 0

function New-XAttribute {
#.Synopsys
#   Creates a new XAttribute (an xml attribute on an XElement for XDocument)
#.Description
#  This is the work-horse for the XML mini-dsl
#.Parameter name
#   The attribute name
#.Parameter value
#  The attribute value
[CmdletBinding()]
Param([Parameter(Mandatory=$true)]$name,[Parameter(Mandatory=$true)]$value)
   New-Object XAttribute $name, $value
}
Set-Alias xa New-XAttribute -EA 0
Set-Alias New-XmlAttribute New-XAttribute -EA 0


function New-XElement {
#.Synopsys
#   Creates a new XElement (an xml tag for XDocument)
#.Description
#  This is the work-horse for the XML mini-dsl
#.Parameter tag
#   The name of the xml tag
#.Parameter args
#   this is where all the dsl magic happens. Please see the Examples. :)
[CmdletBinding()]
Param(
   [Parameter(Mandatory = $true, Position = 0)]
   [System.Xml.Linq.XName]$tag
,
   [AllowNull()][AllowEmptyString()][AllowEmptyCollection()]
   [Parameter(Position=99, Mandatory = $false, ValueFromRemainingArguments=$true)]
   [PSObject[]]$args
)
PROCESS {
  New-Object XElement $(
     $tag
     while($args) {
        $attrib, $value, $args = $args
        if($attrib -is [ScriptBlock]) { # then it's content
           &$attrib
        } elseif ( $value -is [ScriptBlock] -and "-CONTENT".StartsWith($attrib.TrimEnd(':').ToUpper())) { # then it's content
           &$value
        } elseif ( $value -is [XNamespace]) {
           New-Object XAttribute ([XNamespace]::Xmlns + $attrib.TrimStart("-").TrimEnd(':')), $value
           $script:NameSpaceHash.Add($attrib.TrimStart("-").TrimEnd(':'), $value)
        } elseif($value -match "-(?!\d)\w") {
            $args = @($value)+@($args)
        } elseif($value -ne $null) {
           New-Object XAttribute $attrib.TrimStart("-").TrimEnd(':'), $value
        }        
        
     }
   )
}
}
Set-Alias xe New-XElement
Set-Alias New-XmlElement New-XElement

function ConvertFrom-XmlDsl {
Param([ScriptBlock]$script)
   $parserrors = $null
   $global:tokens = [PSParser]::Tokenize( $script, [ref]$parserrors )
   [Array]$duds = $global:tokens | Where-Object { $_.Type -eq "Command" -and !$_.Content.Contains('-') -and ($(Get-Command $_.Content -Type Cmdlet,Function,ExternalScript -EA 0) -eq $Null) }
   [Array]::Reverse( $duds )
   
   [string[]]$ScriptText = "$script" -split "`n"

   ForEach($token in $duds ) {
      # replace : notation with namespace notation
      if( $token.Content.Contains(":") ) {
         $key, $localname = $token.Content -split ":"
         $ScriptText[($token.StartLine - 1)] = $ScriptText[($token.StartLine - 1)].Remove( $token.StartColumn -1, $token.Length ).Insert( $token.StartColumn -1, "'" + $($script:NameSpaceHash[$key] + $localname) + "'" )
      } else {
         $ScriptText[($token.StartLine - 1)] = $ScriptText[($token.StartLine - 1)].Remove( $token.StartColumn -1, $token.Length ).Insert( $token.StartColumn -1, "'" + $($script:NameSpaceHash[''] + $token.Content) + "'" )
      }
      # insert 'xe' before everything (unless it's a valid command)
      $ScriptText[($token.StartLine - 1)] = $ScriptText[($token.StartLine - 1)].Insert( $token.StartColumn -1, "xe " )
   }
   Write-Output ([ScriptBlock]::Create( ($ScriptText -join "`n") ))
}



######## Xaml
#  if($PSVersionTable.CLRVersion -ge "4.0"){
#     trap { continue }
#     [Reflection.Assembly]::LoadWithPartialName("System.Xaml") | Out-Null
#     if("System.Xaml.XamlServices" -as [type]) {
    
   #  }
#  }
   
Export-ModuleMember -alias * -function New-XDocument, New-XAttribute, New-XElement, Remove-XmlNamespace, Get-XmlContent, Set-XmlContent, ConvertTo-Xml, Select-Xml, Update-Xml, Format-Xml, ConvertTo-CliXml, ConvertFrom-CliXml

# SIG # Begin signature block
# MIIdZgYJKoZIhvcNAQcCoIIdVzCCHVMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU7i/4rcSOaFh6MGBAPJLn0Vpy
# SVegghkkMIIDnzCCAoegAwIBAgIQeaKlhfnRFUIT2bg+9raN7TANBgkqhkiG9w0B
# AQUFADBTMQswCQYDVQQGEwJVUzEXMBUGA1UEChMOVmVyaVNpZ24sIEluYy4xKzAp
# BgNVBAMTIlZlcmlTaWduIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EwHhcNMTIw
# NTAxMDAwMDAwWhcNMTIxMjMxMjM1OTU5WjBiMQswCQYDVQQGEwJVUzEdMBsGA1UE
# ChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xNDAyBgNVBAMTK1N5bWFudGVjIFRpbWUg
# U3RhbXBpbmcgU2VydmljZXMgU2lnbmVyIC0gRzMwgZ8wDQYJKoZIhvcNAQEBBQAD
# gY0AMIGJAoGBAKlZZnTaPYp9etj89YBEe/5HahRVTlBHC+zT7c72OPdPabmx8LZ4
# ggqMdhZn4gKttw2livYD/GbT/AgtzLVzWXuJ3DNuZlpeUje0YtGSWTUUi0WsWbJN
# JKKYlGhCcp86aOJri54iLfSYTprGr7PkoKs8KL8j4ddypPIQU2eud69RAgMBAAGj
# geMwgeAwDAYDVR0TAQH/BAIwADAzBgNVHR8ELDAqMCigJqAkhiJodHRwOi8vY3Js
# LnZlcmlzaWduLmNvbS90c3MtY2EuY3JsMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
# MDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AudmVyaXNp
# Z24uY29tMA4GA1UdDwEB/wQEAwIHgDAeBgNVHREEFzAVpBMwETEPMA0GA1UEAxMG
# VFNBMS0zMB0GA1UdDgQWBBS0t/GJSSZg52Xqc67c0zjNv1eSbzANBgkqhkiG9w0B
# AQUFAAOCAQEAHpiqJ7d4tQi1yXJtt9/ADpimNcSIydL2bfFLGvvV+S2ZAJ7R55uL
# 4T+9OYAMZs0HvFyYVKaUuhDRTour9W9lzGcJooB8UugOA9ZresYFGOzIrEJ8Byyn
# PQhm3ADt/ZQdc/JymJOxEdaP747qrPSWUQzQjd8xUk9er32nSnXmTs4rnykr589d
# nwN+bid7I61iKWavkugszr2cf9zNFzxDwgk/dUXHnuTXYH+XxuSqx2n1/M10rCyw
# SMFQTnBWHrU1046+se2svf4M7IV91buFZkQZXZ+T64K6Y57TfGH/yBvZI1h/MKNm
# oTkmXpLDPMs3Mvr1o43c1bCj6SU2VdeB+jCCA8QwggMtoAMCAQICEEe/GZXfjVJG
# Q/fbbUgNMaQwDQYJKoZIhvcNAQEFBQAwgYsxCzAJBgNVBAYTAlpBMRUwEwYDVQQI
# EwxXZXN0ZXJuIENhcGUxFDASBgNVBAcTC0R1cmJhbnZpbGxlMQ8wDQYDVQQKEwZU
# aGF3dGUxHTAbBgNVBAsTFFRoYXd0ZSBDZXJ0aWZpY2F0aW9uMR8wHQYDVQQDExZU
# aGF3dGUgVGltZXN0YW1waW5nIENBMB4XDTAzMTIwNDAwMDAwMFoXDTEzMTIwMzIz
# NTk1OVowUzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMSsw
# KQYDVQQDEyJWZXJpU2lnbiBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAqcqypMzNIK8KfYmsh3XwtE7x38EP
# v2dhvaNkHNq7+cozq4QwiVh+jNtr3TaeD7/R7Hjyd6Z+bzy/k68Numj0bJTKvVIt
# q0g99bbVXV8bAp/6L2sepPejmqYayALhf0xS4w5g7EAcfrkN3j/HtN+HvV96ajEu
# A5mBE6hHIM4xcw1XLc14NDOVEpkSud5oL6rm48KKjCrDiyGHZr2DWFdvdb88qiaH
# XcoQFTyfhOpUwQpuxP7FSt25BxGXInzbPifRHnjsnzHJ8eYiGdvEs0dDmhpfoB6Q
# 5F717nzxfatiAY/1TQve0CJWqJXNroh2ru66DfPkTdmg+2igrhQ7s4fBuwIDAQAB
# o4HbMIHYMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3Au
# dmVyaXNpZ24uY29tMBIGA1UdEwEB/wQIMAYBAf8CAQAwQQYDVR0fBDowODA2oDSg
# MoYwaHR0cDovL2NybC52ZXJpc2lnbi5jb20vVGhhd3RlVGltZXN0YW1waW5nQ0Eu
# Y3JsMBMGA1UdJQQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIBBjAkBgNVHREE
# HTAbpBkwFzEVMBMGA1UEAxMMVFNBMjA0OC0xLTUzMA0GCSqGSIb3DQEBBQUAA4GB
# AEpr+epYwkQcMYl5mSuWv4KsAdYcTM2wilhu3wgpo17IypMT5wRSDe9HJy8AOLDk
# yZNOmtQiYhX3PzchT3AxgPGLOIez6OiXAP7PVZZOJNKpJ056rrdhQfMqzufJ2V7d
# uyuFPrWdtdnhV/++tMV+9c8MnvCX/ivTO1IbGzgn9z9KMIIETzCCA7igAwIBAgIE
# BydYPTANBgkqhkiG9w0BAQUFADB1MQswCQYDVQQGEwJVUzEYMBYGA1UEChMPR1RF
# IENvcnBvcmF0aW9uMScwJQYDVQQLEx5HVEUgQ3liZXJUcnVzdCBTb2x1dGlvbnMs
# IEluYy4xIzAhBgNVBAMTGkdURSBDeWJlclRydXN0IEdsb2JhbCBSb290MB4XDTEw
# MDExMzE5MjAzMloXDTE1MDkzMDE4MTk0N1owbDELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTErMCkG
# A1UEAxMiRGlnaUNlcnQgSGlnaCBBc3N1cmFuY2UgRVYgUm9vdCBDQTCCASIwDQYJ
# KoZIhvcNAQEBBQADggEPADCCAQoCggEBAMbM5XPm+9S75S0tMqbf5YE/yc0lSbZx
# KsPVlDRnogocsF9ppkCxxLeyj9CYpKlBWTrT3JTWPNt0OKRKzE0lgvdKpVMSOO7z
# SW1xkX5jtqumX8OkhPhPYlG++MXs2ziS4wblCJEMxChBVfvLWokVfnHoNb9Ncgk9
# vjo4UFt3MRuNs8ckRZqnrG0AFFoEt7oT61EKmEFBIk5lYYeBQVCmeVyJ3hlKV9Uu
# 5l0cUyx+mM0aBhakaHPQNAQTXKFx01p8VdteZOE3hzBWBOURtCmAEvF5OYiiAhF8
# J2a3iLd48soKqDirCmTCv2ZdlYTBoSUeh10aUAsgEsxBu24LUTi4S8sCAwEAAaOC
# AW8wggFrMBIGA1UdEwEB/wQIMAYBAf8CAQEwUwYDVR0gBEwwSjBIBgkrBgEEAbE+
# AQAwOzA5BggrBgEFBQcCARYtaHR0cDovL2N5YmVydHJ1c3Qub21uaXJvb3QuY29t
# L3JlcG9zaXRvcnkuY2ZtMA4GA1UdDwEB/wQEAwIBBjCBiQYDVR0jBIGBMH+heaR3
# MHUxCzAJBgNVBAYTAlVTMRgwFgYDVQQKEw9HVEUgQ29ycG9yYXRpb24xJzAlBgNV
# BAsTHkdURSBDeWJlclRydXN0IFNvbHV0aW9ucywgSW5jLjEjMCEGA1UEAxMaR1RF
# IEN5YmVyVHJ1c3QgR2xvYmFsIFJvb3SCAgGlMEUGA1UdHwQ+MDwwOqA4oDaGNGh0
# dHA6Ly93d3cucHVibGljLXRydXN0LmNvbS9jZ2ktYmluL0NSTC8yMDE4L2NkcC5j
# cmwwHQYDVR0OBBYEFLE+w2kD+L9HAdSYJhoIAu9jZCvDMA0GCSqGSIb3DQEBBQUA
# A4GBAC52hdk3lm2vifMGeIIxxEYHH2XJjrPJVHjm0ULfdS4eVer3+psEwHV70Xk8
# Bex5xFLdpgPXp1CZPwVZ2sZV9IacDWejSQSVMh3Hh+yFr2Ru1cVfCadAfRa6SQ2i
# /fbfVTBs13jGuc9YKWQWTKMggUexRJKEFhtvSrwhxgo97TPKMIIGnzCCBYegAwIB
# AgIQDmkGmMIUyHq1tgS5FjzRkDANBgkqhkiG9w0BAQUFADBzMQswCQYDVQQGEwJV
# UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQu
# Y29tMTIwMAYDVQQDEylEaWdpQ2VydCBIaWdoIEFzc3VyYW5jZSBDb2RlIFNpZ25p
# bmcgQ0EtMTAeFw0xMjAzMjAwMDAwMDBaFw0xMzAzMjIxMjAwMDBaMG0xCzAJBgNV
# BAYTAlVTMREwDwYDVQQIEwhOZXcgWW9yazEXMBUGA1UEBxMOV2VzdCBIZW5yaWV0
# dGExGDAWBgNVBAoTD0pvZWwgSC4gQmVubmV0dDEYMBYGA1UEAxMPSm9lbCBILiBC
# ZW5uZXR0MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA2ogGAG89d1jM
# fRJv2d3U1lCsW8ok7GkjnLYDn0zC1ALq11rWN5NVwVbn133i+KV0O8kM5vd2M7xE
# 8CnVAgybjkrvRD2IqMtp4SrwQuiGiVGsNVWO3vSLHcWsS/I7N0UIpS5PhTuFB4Pc
# Oy/MHR4F2g6JLMrAtkpYWxauAFZfFwuEfm6vqWobHTDt5wG+zqOTxMSi1UvL5fEM
# DoejGqqriIx5mKDzrvUb/ALNKZ1rGPWlT7O0/UHrV5VuOfgij4XVKBAdcg9JLxky
# AEIJ+VvVQ2Jn3lVONCCHbfu5IVhddMru81U/v5Wrj80Zrwh2TH25qlclUKr6eXRL
# tP+xFm23CwIDAQABo4IDMzCCAy8wHwYDVR0jBBgwFoAUl0gD6xUIa7myWCPMlC7x
# xmXSZI4wHQYDVR0OBBYEFJicRKq/XsBWRuKzU6eTUCBCCU65MA4GA1UdDwEB/wQE
# AwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzBpBgNVHR8EYjBgMC6gLKAqhihodHRw
# Oi8vY3JsMy5kaWdpY2VydC5jb20vaGEtY3MtMjAxMWEuY3JsMC6gLKAqhihodHRw
# Oi8vY3JsNC5kaWdpY2VydC5jb20vaGEtY3MtMjAxMWEuY3JsMIIBxAYDVR0gBIIB
# uzCCAbcwggGzBglghkgBhv1sAwEwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3
# LmRpZ2ljZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUH
# AgIwggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQBy
# AHQAaQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBj
# AGUAcAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAg
# AEMAUAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQ
# AGEAcgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBt
# AGkAdAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBj
# AG8AcgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBl
# AHIAZQBuAGMAZQAuMIGGBggrBgEFBQcBAQR6MHgwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBQBggrBgEFBQcwAoZEaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0SGlnaEFzc3VyYW5jZUNvZGVTaWduaW5nQ0Et
# MS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQUFAAOCAQEAHIfeYpO0Jtdi
# /TpcI6eWQIYU2ALO847Q91jLE6WiU6u8wN6tkHqgeOls070SDUK+C1rVoXKKZ0Je
# c2k1dYukKPkyf3qURPyh/aC3hJ0Wwbje7fK79Lt9ZHwJORpesJrwa8T63l3qLLLl
# PaIYo/bqiMpNZRfOclukKg2hO67yMaQl8DEL/D5UP1XZShF2zbauH627zEC5KXGZ
# Y2yUbmWG2N0oHxr+q4Gyfd0MPtU5avWOILB0ZsN+br+SCVVK6nKzauXMk4HXmKHa
# X7cysqpmQiFb7/J7tPQ037KQKHCY/Z+fl0arRCiHih/Q/5owv51WSKPiaUrkBvdJ
# 0mKVK+McHzCCBr8wggWnoAMCAQICEAgcV+5dcOuboLFSDHKcGwkwDQYJKoZIhvcN
# AQEFBQAwbDELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTErMCkGA1UEAxMiRGlnaUNlcnQgSGlnaCBB
# c3N1cmFuY2UgRVYgUm9vdCBDQTAeFw0xMTAyMTAxMjAwMDBaFw0yNjAyMTAxMjAw
# MDBaMHMxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNV
# BAsTEHd3dy5kaWdpY2VydC5jb20xMjAwBgNVBAMTKURpZ2lDZXJ0IEhpZ2ggQXNz
# dXJhbmNlIENvZGUgU2lnbmluZyBDQS0xMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8A
# MIIBCgKCAQEAxfkj5pQnxIAUpIAyX0CjjW9wwOU2cXE6daSqGpKUiV6sI3HLTmd9
# QT+q40u3e76dwag4j2kvOiTpd1kSx2YEQ8INJoKJQBnyLOrnTOd8BRq4/4gJTyY3
# 7zqk+iJsiMlKG2HyrhBeb7zReZtZGGDl7im1AyqkzvGDGU9pBXMoCfsiEJMioJAZ
# Gkwx8tMr2IRDrzxj/5jbINIJK1TB6v1qg+cQoxJx9dbX4RJ61eBWWs7qAVtoZVvB
# P1hSM6k1YU4iy4HKNqMSywbWzxtNGH65krkSz0Am2Jo2hbMVqkeThGsHu7zVs94l
# ABGJAGjBKTzqPi3uUKvXHDAGeDylECNnkQIDAQABo4IDVDCCA1AwDgYDVR0PAQH/
# BAQDAgEGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMIIBwwYDVR0gBIIBujCCAbYwggGy
# BghghkgBhv1sAzCCAaQwOgYIKwYBBQUHAgEWLmh0dHA6Ly93d3cuZGlnaWNlcnQu
# Y29tL3NzbC1jcHMtcmVwb3NpdG9yeS5odG0wggFkBggrBgEFBQcCAjCCAVYeggFS
# AEEAbgB5ACAAdQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBj
# AGEAdABlACAAYwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBu
# AGMAZQAgAG8AZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAARQBWACAAQwBQ
# AFMAIABhAG4AZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAg
# AEEAZwByAGUAZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABp
# AGEAYgBpAGwAaQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwBy
# AGEAdABlAGQAIABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBl
# AC4wDwYDVR0TAQH/BAUwAwEB/zB/BggrBgEFBQcBAQRzMHEwJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBJBggrBgEFBQcwAoY9aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0SGlnaEFzc3VyYW5jZUVWUm9vdENB
# LmNydDCBjwYDVR0fBIGHMIGEMECgPqA8hjpodHRwOi8vY3JsMy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRIaWdoQXNzdXJhbmNlRVZSb290Q0EuY3JsMECgPqA8hjpodHRw
# Oi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRIaWdoQXNzdXJhbmNlRVZSb290
# Q0EuY3JsMB0GA1UdDgQWBBSXSAPrFQhrubJYI8yULvHGZdJkjjAfBgNVHSMEGDAW
# gBSxPsNpA/i/RwHUmCYaCALvY2QrwzANBgkqhkiG9w0BAQUFAAOCAQEAggXpha+n
# TL+vzj2y6mCxaN5nwtLLJuDDL5u1aw5TkIX2m+A1Av/6aYOqtHQyFDwuEEwomwqt
# CAn584QRk4/LYEBW6XcvabKDmVWrRySWy39LsBC0l7/EpZkG/o7sFFAeXleXy0e5
# NNn8OqL/UCnCCmIE7t6WOm+gwoUPb/wI5DJ704SuaWAJRiac6PD//4bZyAk6ZsOn
# No8YT+ixlpIuTr4LpzOQrrxuT/F+jbRGDmT5WQYiIWQAS+J6CAPnvImQnkJPAcC2
# Fn916kaypVQvjJPNETY0aihXzJQ/6XzIGAMDBH5D2vmXoVlH2hKq4G04AF01K8Ui
# hssGyrx6TT0mRjGCA6wwggOoAgEBMIGHMHMxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMjAwBgNV
# BAMTKURpZ2lDZXJ0IEhpZ2ggQXNzdXJhbmNlIENvZGUgU2lnbmluZyBDQS0xAhAO
# aQaYwhTIerW2BLkWPNGQMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQi/t811+xfM0LNj5x8Yh4W
# 4VVTRTANBgkqhkiG9w0BAQEFAASCAQA216Utho6PZL213Vbd+P5nXS9mRBz581vL
# EkhYMvfduQeikLHIIP2CqEGcnwuwNWTzHlN0UkdsDtP9tLdXQt+czvhikd6rzrKH
# VcrnfE4IPQt6QT+06PAb4Dwt2RVV9dpstHhKZIWn+GoFuLHuJQzPMgdulYeKMFO2
# ZLrEkmF5+DWzMpBC80WLlpwDJOoUmiuK4oQfzMzqFLBxpw5RAQcA+AA3ohWtQPuF
# c7ywquc1c8dCxG5cYaRtFo+LLf6jx4FO24vreKWLbEOx0cTb8PNWvrT34lgnTgUI
# pWLJkSRGzywnq3w+keexYi4GC16PfJKnVV8qsRri/wG/AdBQ+ovnoYIBfzCCAXsG
# CSqGSIb3DQEJBjGCAWwwggFoAgEBMGcwUzELMAkGA1UEBhMCVVMxFzAVBgNVBAoT
# DlZlcmlTaWduLCBJbmMuMSswKQYDVQQDEyJWZXJpU2lnbiBUaW1lIFN0YW1waW5n
# IFNlcnZpY2VzIENBAhB5oqWF+dEVQhPZuD72to3tMAkGBSsOAwIaBQCgXTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xMjA2MjEwNTAw
# NTlaMCMGCSqGSIb3DQEJBDEWBBSODCUreYZBSE+ys+civTtfLRjcLjANBgkqhkiG
# 9w0BAQEFAASBgFrvbqX/d1ys5yFuOrI2PeQOPEAkE5DdaU7KmGtYUR7B8aLQGXfw
# EP3AoFVdswv8zHnZuMMiV3xhtzAT0w0xQKj+CQQ7iQLgyWpSU0zonMT/bRkDglKF
# JfC6tpo0dJnqqLXfc3p9JJiSJyI5HneFwtbVSyhmWfYLBJ6kB6Ofypn+
# SIG # End signature block