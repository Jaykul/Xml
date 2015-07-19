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
&{ 
$local:xlr8r = [psobject].assembly.gettype("System.Management.Automation.TypeAccelerators")
$local:xlinq = [Reflection.Assembly]::Load("System.Xml.Linq, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
$xlinq.GetTypes() | ? { $_.IsPublic -and !$_.IsSerializable -and $_.Name -ne "Extensions" -and !$xlr8r::Get[$_.Name] } | % {
  $xlr8r::Add( $_.Name, $_.FullName )
}

if(!$xlr8r::Get["Dictionary"]) {
   $xlr8r::Add( "Dictionary", "System.Collections.Generic.Dictionary``2, mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" )
}
if(!$xlr8r::Get["PSParser"]) {
   $xlr8r::Add( "PSParser", "System.Management.Automation.PSParser, System.Management.Automation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" )
}
}


function Import-Xml {
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
Set-Alias ipxml Import-Xml
Set-Alias ipx Import-Xml
Set-Alias Get-Xml Import-Xml
Set-Alias gxml Import-Xml
Set-Alias gx Import-Xml

function Export-Xml {
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
Set-Alias epxml Export-Xml
Set-Alias epx Export-Xml
Set-Alias Set-Xml Export-Xml
Set-Alias sxml Export-Xml
Set-Alias sx Export-Xml

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
      $type = [type]::gettype("System.Management.Automation.Deserializer")
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
            write-warning "Could not deserialize $string: $_"
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
      $type = [type]::gettype("System.Management.Automation.Serializer")
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
   
Export-ModuleMember -alias * -function New-XDocument, New-XAttribute, New-XElement, Remove-XmlNamespace, Import-Xml, Export-Xml, ConvertTo-Xml, Select-Xml, Update-Xml, Format-Xml, ConvertTo-CliXml, ConvertFrom-CliXml

# SIG # Begin signature block
# MIIIDQYJKoZIhvcNAQcCoIIH/jCCB/oCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUiMRLI7iLPeWuXLdb7lMdZlym
# fzCgggUrMIIFJzCCBA+gAwIBAgIQHCAgf57pVOnJcjKrMO/dtjANBgkqhkiG9w0B
# AQUFADCBlTELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAlVUMRcwFQYDVQQHEw5TYWx0
# IExha2UgQ2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMSEwHwYD
# VQQLExhodHRwOi8vd3d3LnVzZXJ0cnVzdC5jb20xHTAbBgNVBAMTFFVUTi1VU0VS
# Rmlyc3QtT2JqZWN0MB4XDTExMDQyNTAwMDAwMFoXDTEyMDQyNDIzNTk1OVowgZUx
# CzAJBgNVBAYTAlVTMQ4wDAYDVQQRDAUwNjg1MDEUMBIGA1UECAwLQ29ubmVjdGlj
# dXQxEDAOBgNVBAcMB05vcndhbGsxFjAUBgNVBAkMDTQ1IEdsb3ZlciBBdmUxGjAY
# BgNVBAoMEVhlcm94IENvcnBvcmF0aW9uMRowGAYDVQQDDBFYZXJveCBDb3Jwb3Jh
# dGlvbjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANaXR+8W+aH6ofiO
# bZRdIRBuvemJ/8c2fDwbHVLBMieiG9Eqs5+XKZ3M17Sz8GBNzQ4bluk2esIycr9z
# yR/ISBjVxz1RcxH79vuvM6husOAKc2YhnGqA6vmfWokmEfDrOH1qLKA4226tPXBE
# eNTSDrYtXFZ6jYWv9kqGcRMBzV7NPvJwQoMDEl1dbNAXo99RaHGjAfVkCSNYMM11
# jzZ2/DyAqVgKVnNviRQ+Wq8HPxP7Eqg/6b2DVw1Nokg3IDeyFRlo2he09YwVEV+r
# GLvjUBmVRQPauJIr1EUgz85byWtYAUWOXNIFiWrqOKj/Clvi5Y9M05a1TwSi4o0F
# yfa4keECAwEAAaOCAW8wggFrMB8GA1UdIwQYMBaAFNrtZHQUnBQ8q92Zqb1bKE2L
# PMnYMB0GA1UdDgQWBBTKaDgQ0lToUHAI+jy/CDn0BluXFjAOBgNVHQ8BAf8EBAMC
# B4AwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDAzARBglghkgBhvhC
# AQEEBAMCBBAwRgYDVR0gBD8wPTA7BgwrBgEEAbIxAQIBAwIwKzApBggrBgEFBQcC
# ARYdaHR0cHM6Ly9zZWN1cmUuY29tb2RvLm5ldC9DUFMwQgYDVR0fBDswOTA3oDWg
# M4YxaHR0cDovL2NybC51c2VydHJ1c3QuY29tL1VUTi1VU0VSRmlyc3QtT2JqZWN0
# LmNybDA0BggrBgEFBQcBAQQoMCYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmNv
# bW9kb2NhLmNvbTAhBgNVHREEGjAYgRZKb2VsLkJlbm5ldHRAeGVyb3guY29tMA0G
# CSqGSIb3DQEBBQUAA4IBAQAzwUwy00sEOggAavqrNoNeEVr0o623DgG2/2EuTsA6
# 2wI0Arb5D0s/icanshHgWwJZBEMZeHa17Ai/E3foCpj6rA3Y4vIQXHukluiSmjU6
# bWTgF5VbNTpvhlOO6E7Ya/rBr4oj4dqTEErkS7acgBHKrjPOptCiU4BSDqtl0k5z
# OIiawyRSITHYEcCcI0Yl7VIz8EDblOQI3b4JGYcmJ7D+peYrnI2zoQyXDigcIj4l
# VlipnjnvYsF+JbPkQY8XbMO+Yc490Bh8BMXPtuLR1KMuIXPK7DKX7JPmJcY7kKF/
# SPviyk0HE7Rldsses73UF8wT3lgj57FiUqX8FdTa7NllMYICTDCCAkgCAQEwgaow
# gZUxCzAJBgNVBAYTAlVTMQswCQYDVQQIEwJVVDEXMBUGA1UEBxMOU2FsdCBMYWtl
# IENpdHkxHjAcBgNVBAoTFVRoZSBVU0VSVFJVU1QgTmV0d29yazEhMB8GA1UECxMY
# aHR0cDovL3d3dy51c2VydHJ1c3QuY29tMR0wGwYDVQQDExRVVE4tVVNFUkZpcnN0
# LU9iamVjdAIQHCAgf57pVOnJcjKrMO/dtjAJBgUrDgMCGgUAoHgwGAYKKwYBBAGC
# NwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUgc2AkVqE
# EN849j8G0GohLlwueJ4wDQYJKoZIhvcNAQEBBQAEggEALsDFBvLn1SEBk+fyLgHL
# ibkAHDtrWMwfSLvOv3Sw61Oa+mviJp8H0pODa+3zfJxEoRa0JHYWUSMutNG0CIg9
# YiEf1HLNoYam9s9rtR9YWdv9wx3CkU9Ci2yGfB4oEc1eCg/1prEHG7XWjqZDC6uG
# vhFUNDHp9sWSF/ruVS3RBEm2VBJkvehpIsB4WsQ0aunDhsU4t2oFnFYC0leaasFy
# rLNzzj9T4CZFydtTVh5xvXG0lWv5SMNhqegYjl+gfxGxrLBrNZiCTkC8vX2c5EHN
# AhsqD53tH0wpC+icLoK+f/DPzi6xqJVOaQ1oruFk94RYNRAFGcgWInqKc2Gvsf6P
# rA==
# SIG # End signature block