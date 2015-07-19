#requires -version 2.0

# Improves over the built-in Select-XML by leveraging Remove-XmlNamespace http://poshcode.org/1492 
# to provide a -RemoveNamespace parameter -- if it's supplied, all of the namespace declarations 
# and prefixes are removed from all XML nodes (by an XSL transform) before searching. 
# IMPORTANT: returned results *will not* have namespaces in them, even if the input XML did. 

# Also, only raw XmlNodes are returned from this function, so the output isn't completely compatible 
# with the built in Select-Xml. It's equivalent to using Select-Xml ... | Select-Object -Expand Node

# Version History:
# Select-Xml 2.0 was the first script version I wrote, and it didn't function identically to the built-in Select-Xml with regards to parameter parsing
# Select-Xml 2.1 matched the built-in Select-Xml parameter sets, it's now a drop-in replacement if you were using the original with: Select-Xml ... | Select-Object -Expand Node
# Select-Xml 2.2 fixes a bug in the -Content parameterset where -RemoveNamespace was *presumed*
# Added New-Xml and associated generation functions for my XML DSL


$xlr8r = [type]::gettype("System.Management.Automation.TypeAccelerators")
$xlinq = [Reflection.Assembly]::Load("System.Xml.Linq, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
$xlinq.GetTypes() | ? { $_.IsPublic -and !$_.IsSerializable -and $_.Name -ne "Extensions" -and !$xlr8r::Get[$_.Name] } | % {
  $xlr8r::Add( $_.Name, $_.FullName )
}

function Format-XML {
#.Synopsis
#   Pretty-print formatted XML source
#.Description
#	Runs an XmlDocument through an auto-indenting XmlWriter
#.Parameter Xml
#	The Xml Document
#.Parameter Indent
#	The indent level (defaults to 2 spaces)
Param(
	[Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
	[xml]$xml
,
	[Parameter(Mandatory=$false)]
	$indent=2
)
    $StringWriter = New-Object System.IO.StringWriter
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter
    $xmlWriter.Formatting = "indented"
    $xmlWriter.Indentation = $Indent
    $xml.WriteContentTo($XmlWriter)
    $XmlWriter.Flush()
    $StringWriter.Flush()
    Write-Output $StringWriter.ToString()
}
Set-Alias fxml Format-Xml -Option AllScope

function Select-Xml {
#.Synopsis
#  The Select-XML cmdlet lets you use XPath queries to search for text in XML strings and documents. Enter an XPath query, and use the Content, Path, or Xml parameter to specify the XML to be searched.
#.Description
#  Improves over the built-in Select-XML by leveraging Remove-XmlNamespace to provide a -RemoveNamespace parameter -- if it's supplied, all of the namespace declarations and prefixes are removed from all XML nodes (by an XSL transform) before searching.  
#  
#  However, only raw XmlNodes are returned from this function, so the output isn't currently compatible with the built in Select-Xml, but is equivalent to using Select-Xml ... | Select-Object -Expand Node
#
#  Also note that if the -RemoveNamespace switch is supplied the returned results *will not* have namespaces in them, even if the input XML did, and entities get expanded automatically.
#.Parameter Content
#  Specifies a string that contains the XML to search. You can also pipe strings to Select-XML.
#.Parameter Namespace
#   Specifies a hash table of the namespaces used in the XML. Use the format @{<namespaceName> = <namespaceUri>}.
#.Parameter Path
#   Specifies the path and file names of the XML files to search.  Wildcards are permitted.
#.Parameter Xml
#  Specifies one or more XML nodes to search.
#.Parameter XPath
#  Specifies an XPath search query. The query language is case-sensitive. This parameter is required.
#.Parameter RemoveNamespace
#  Allows the execution of XPath queries without namespace qualifiers. 
#  
#  If you specify the -RemoveNamespace switch, all namespace declarations and prefixes are actually removed from the Xml before the XPath search query is evaluated, and your XPath query should therefore NOT contain any namespace prefixes.
# 
#  Note that this means that the returned results *will not* have namespaces in them, even if the input XML did, and entities get expanded automatically.
[CmdletBinding(DefaultParameterSetName="Xml")]
PARAM(
   [Parameter(Position=1,ParameterSetName="Path",Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
   [ValidateNotNullOrEmpty()]
   [Alias("PSPath")]
   [String[]]$Path
,
   [Parameter(Position=1,ParameterSetName="Xml",Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
   [ValidateNotNullOrEmpty()]
   [Alias("Node")]
   [System.Xml.XmlNode[]]$Xml
,
   [Parameter(ParameterSetName="Content",Mandatory=$true,ValueFromPipeline=$true)]
   [ValidateNotNullOrEmpty()]
   [String[]]$Content
,
   [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false)]
   [ValidateNotNullOrEmpty()]
   [Alias("Query")]
   [String[]]$XPath
,
   [Parameter(Mandatory=$false)]
   [ValidateNotNullOrEmpty()]
   [Hashtable]$Namespace
,
   [Switch]$RemoveNamespace
)
BEGIN {
   function Select-Node {
   PARAM([Xml.XmlNode]$Xml, [String[]]$XPath, $NamespaceManager)
   BEGIN {
      foreach($node in $xml) {
         if($NamespaceManager -is [Hashtable]) {
            $nsManager = new-object System.Xml.XmlNamespaceManager $node.NameTable
            foreach($ns in $Namespace.GetEnumerator()) {
               $nsManager.AddNamespace( $ns.Key, $ns.Value )
            }
         }
         
         foreach($path in $xpath) {
            $node.SelectNodes($path, $NamespaceManager)
   }  }  }  }

   [Text.StringBuilder]$XmlContent = [String]::Empty
}

PROCESS {
   $NSM = $Null; if($PSBoundParameters.ContainsKey("Namespace")) { $NSM = $Namespace }

   switch($PSCmdlet.ParameterSetName) {
      "Content" {
         $null = $XmlContent.AppendLine( $Content -Join "`n" )
      }
      "Path" {
         foreach($file in Get-ChildItem $Path) {
            [Xml]$Xml = Get-Content $file
            if($RemoveNamespace) {
               $Xml = Remove-XmlNamespace $Xml
            }
            Select-Node $Xml $XPath  $NSM
         }
      }
      "Xml" {
         foreach($node in $Xml) {
            if($RemoveNamespace) {
               $node = Remove-XmlNamespace $node
            }
            Select-Node $node $XPath $NSM
         }
      }
   }
}
END {
   if($PSCmdlet.ParameterSetName -eq "Content") {
      [Xml]$Xml = $XmlContent.ToString()
      if($RemoveNamespace) {
         $Xml = Remove-XmlNamespace $Xml
      }
      Select-Node $Xml $XPath  $NSM
   }
}

}
Set-Alias slxml Select-Xml -Option AllScope

function Convert-Node {
#.Synopsis 
# Convert a single XML Node via XSL stylesheets
param(
[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
[System.Xml.XmlReader]$XmlReader,
[Parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false)]
[System.Xml.Xsl.XslCompiledTransform]$StyleSheet
) 
PROCESS {
   $output = New-Object IO.StringWriter
   $StyleSheet.Transform( $XmlReader, $null, $output )
   Write-Output $output.ToString()
}
}
   
function Convert-Xml {
#.Synopsis
#  The Convert-XML function lets you use Xslt to transform XML strings and documents.
#.Description
#.Parameter Content
#  Specifies a string that contains the XML to search. You can also pipe strings to Select-XML.
#.Parameter Namespace
#   Specifies a hash table of the namespaces used in the XML. Use the format @{<namespaceName> = <namespaceUri>}.
#.Parameter Path
#   Specifies the path and file names of the XML files to search.  Wildcards are permitted.
#.Parameter Xml
#  Specifies one or more XML nodes to search.
#.Parameter Xsl
#  Specifies an Xml StyleSheet to transform with...
[CmdletBinding(DefaultParameterSetName="Xml")]
PARAM(
   [Parameter(Position=1,ParameterSetName="Path",Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
   [ValidateNotNullOrEmpty()]
   [Alias("PSPath")]
   [String[]]$Path
,
   [Parameter(Position=1,ParameterSetName="Xml",Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
   [ValidateNotNullOrEmpty()]
   [Alias("Node")]
   [System.Xml.XmlNode[]]$Xml
,
   [Parameter(ParameterSetName="Content",Mandatory=$true,ValueFromPipeline=$true)]
   [ValidateNotNullOrEmpty()]
   [String[]]$Content
,
   [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false)]
   [ValidateNotNullOrEmpty()]
   [Alias("StyleSheet")]
   [String[]]$Xslt
)
BEGIN { 
   $StyleSheet = New-Object System.Xml.Xsl.XslCompiledTransform
   if(Test-Path @($Xslt)[0] -EA 0) { 
      Write-Verbose "Loading Stylesheet from $(Resolve-Path @($Xslt)[0])"
      $StyleSheet.Load( (Resolve-Path @($Xslt)[0]) )
   } else {
      Write-Verbose "$Xslt"
      $StyleSheet.Load(([System.Xml.XmlReader]::Create((New-Object System.IO.StringReader ($Xslt -join "`n")))))
   }
   [Text.StringBuilder]$XmlContent = [String]::Empty 
}
PROCESS {
   switch($PSCmdlet.ParameterSetName) {
      "Content" {
         $null = $XmlContent.AppendLine( $Content -Join "`n" )
      }
      "Path" {
         foreach($file in Get-ChildItem $Path) {
            Convert-Node -Xml ([System.Xml.XmlReader]::Create((Resolve-Path $file))) $StyleSheet
         }
      }
      "Xml" {
         foreach($node in $Xml) {
            Convert-Node -Xml (New-Object Xml.XmlNodeReader $node) $StyleSheet
         }
      }
   }
}
END {
   if($PSCmdlet.ParameterSetName -eq "Content") {
      [Xml]$Xml = $XmlContent.ToString()
      Convert-Node -Xml $Xml $StyleSheet
   }
}
}
Set-Alias cvxml Convert-Xml -Option AllScope

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
   [Parameter(Position=1,ParameterSetName="Path",Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
   [ValidateNotNullOrEmpty()]
   [Alias("PSPath")]
   [String[]]$Path
,
   [Parameter(Position=1,ParameterSetName="Xml",Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
   [ValidateNotNullOrEmpty()]
   [Alias("Node")]
   [System.Xml.XmlNode[]]$Xml
,
   [Parameter(ParameterSetName="Content",Mandatory=$true,ValueFromPipeline=$true)]
   [ValidateNotNullOrEmpty()]
   [String[]]$Content
,
   [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false)]
   [ValidateNotNullOrEmpty()]
   [Alias("StyleSheet")]
   [String[]]$Xslt
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
   switch($PSCmdlet.ParameterSetName) {
      "Content" {
         $null = $XmlContent.AppendLine( $Content -Join "`n" )
      }
      "Path" {
         foreach($file in Get-ChildItem $Path) {
            [Xml]$Xml = Get-Content $file
            Convert-Node -Xml $Xml $StyleSheet
         }
      }
      "Xml" {
         $Xml | Convert-Node $StyleSheet
      }
   }
}
END {
   if($PSCmdlet.ParameterSetName -eq "Content") {
      [Xml]$Xml = $XmlContent.ToString()
      Convert-Node -Xml $Xml $StyleSheet
   }
}
}
Set-Alias rmns Remove-XmlNamespace -Option AllScope

function New-XDocument {
#.Synopsys
#	Creates a new XDocument (the new xml document type)
#.Description
#  This is the root for a new XML mini-dsl, akin to New-BootsWindow for XAML
#  It creates a new XDocument, and takes scritpblock(s) to define it's contents
#.Parameter root
#	The root node name
#.Parameter version
#	Optional: the XML version. Defaults to 1.0
#.Parameter encoding
#	Optional: the Encoding. Defaults to UTF-8
#.Parameter standalone
#  Optional: whether to specify standalone in the xml declaration. Defaults to "yes"
#.Parameter args
#	this is where all the dsl magic happens. Please see the Examples. :)
#
#.Example
# [XNamespace]$dc = "http`://purl.org/dc/elements/1.1"
# $xml = New-XDocument rss -dc $dc -version "2.0" {
#    xe channel {
#       xe title {"Test RSS Feed"}
#       xe link {"http`://HuddledMasses.org"}
#       xe description {"An RSS Feed generated simply to demonstrate my XML DSL"}
#       xe ($dc + "language") {"en"}
#       xe ($dc + "creator") {"Jaykul@HuddledMasses.org"}
#       xe ($dc + "rights") {"Copyright 2009, CC-BY"}
#       xe ($dc + "date") {(Get-Date -f u) -replace " ","T"}
#       xe item {
#          xe title {"The First Item"}
#          xe link {"http`://huddledmasses.org/new-site-new-layout-lost-posts/"}
#          xe guid -isPermaLink true {"http`://huddledmasses.org/new-site-new-layout-lost-posts/"}
#          xe description {"Ema Lazarus' Poem"}
#          xe pubDate  {(Get-Date 10/31/2003 -f u) -replace " ","T"}
#       }
#    }
# }
#
# $xml.Declaration.ToString()  ## I can't find a way to have this included in the $xml.ToString()
# $xml.ToString()
#
#
# OUTPUT: (NOTE: I added the space in the http: to paste it on PoshCode -- those aren't in the output)
#
#
# <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
# <rss xmlns:dc="http ://purl.org/dc/elements/1.1" version="2.0">
#   <channel>
#     <title>Test RSS Feed</title>
#     <link>http ://HuddledMasses.org</link>
#     <description>An RSS Feed generated simply to demonstrate my XML DSL</description>
#     <dc:language>en</dc:language>
#     <dc:creator>Jaykul@HuddledMasses.org</dc:creator>
#     <dc:rights>Copyright 2009, CC-BY</dc:rights>
#     <dc:date>2009-07-26T00:50:08Z</dc:date>
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
#.Example 
#  This time with a default namespace
## IMPORTANT! ## NOTE that I use the "xe" shortcut which is redefined when you specify a namespace
##            ## for the root element, so that all child elements (by default) inherit that.
##            ## You can still control the prefixes by passing the namespace as a parameter
##            ## e.g.: -atom $atom
## The `: in the http`: is still only there for PoshCode, you can just use http: ...
#
#   [XNamespace]$atom="http`://www.w3.org/2005/Atom"
#   [XNamespace]$dc = "http`://purl.org/dc/elements/1.1"
#  
#   New-Xml ($atom + "feed") -Encoding "UTF-16" -$([XNamespace]::Xml +'lang') "en-US" -dc $dc {
#      xe title {"Test First Entry"}
#      xe link {"http`://HuddledMasses.org"}
#      xe updated {(Get-Date -f u) -replace " ","T"}
#      xe author {
#         xe name {"Joel Bennett"}
#         xe uri {"http`://HuddledMasses.org"}
#      }
#      xe id {"http`://huddledmasses.org/" }
#
#      xe entry {
#         xe title {"Test First Entry"}
#         xe link {"http`://HuddledMasses.org/new-site-new-layout-lost-posts/" }
#         xe id {"http`://huddledmasses.org/new-site-new-layout-lost-posts/" }
#         xe updated {(Get-Date 10/31/2003 -f u) -replace " ","T"}
#         xe summary {"Ema Lazarus' Poem"}
#         xe link -rel license -href "http://creativecommons.org/licenses/by/3.0/" -title "CC By-Attribution"
#         xe ($dc + "rights") {"Copyright 2009, Some rights reserved (licensed under the Creative Commons Attribution 3.0 Unported license)"}
#         xe category -scheme "http://huddledmasses.org/tag/" -term "huddled-masses"
#      }
#   } | % { $_.Declaration.ToString(); $_.ToString() }
#
#
#  OUTPUT: (NOTE: I added the spaces again to the http: to paste it on PoshCode)
#
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
   [Parameter(Position=99, Mandatory = $false, ValueFromRemainingArguments=$true)]
   [PSObject[]]$args
)
BEGIN {
   if(![string]::IsNullOrEmpty( $root.NamespaceName )) {
      Function New-QualifiedXElement {
         Param([System.Xml.Linq.XName]$tag)
         if([string]::IsNullOrEmpty( $tag.NamespaceName )) {
            $tag = $($root.Namespace) + $tag
         }
         New-XElement $tag @args
      }
      Set-Alias xe New-QualifiedXElement
   }
}
PROCESS {
   #New-Object XDocument (New-Object XDeclaration "1.0", "UTF-8", "yes"),(
   New-Object XDocument (New-Object XDeclaration $Version, $Encoding, $standalone),(
      New-Object XElement $(
         $root
         while($args) {
            $attrib, $value, $args = $args
            if($attrib -is [ScriptBlock]) {
               &$attrib
            } elseif ( $value -is [ScriptBlock] -and "-Content".StartsWith($attrib)) {
               &$value
            } elseif ( $value -is [XNamespace]) {
               New-XAttribute ([XNamespace]::Xmlns + $attrib.TrimStart("-")) $value
            } else {
               New-XAttribute $attrib.TrimStart("-") $value
            }
         }
      ))
}
END {
   Set-Alias xe New-XElement
}
}

Set-Alias xml New-XDocument -Option AllScope
Set-Alias New-Xml New-XDocument -Option AllScope

function New-XAttribute {
#.Synopsys
#	Creates a new XAttribute (an xml attribute on an XElement for XDocument)
#.Description
#  This is the work-horse for the XML mini-dsl
#.Parameter name
#	The attribute name
#.Parameter value
#  The attribute value
Param([Parameter(Mandatory=$true)]$name,[Parameter(Mandatory=$true)]$value)
   New-Object XAttribute $name,$value
}
Set-Alias xa New-XAttribute -Option AllScope
Set-Alias New-XmlAttribute New-XAttribute -Option AllScope


function New-XElement {
#.Synopsys
#	Creates a new XElement (an xml tag for XDocument)
#.Description
#  This is the work-horse for the XML mini-dsl
#.Parameter tag
#	The name of the xml tag
#.Parameter args
#	this is where all the dsl magic happens. Please see the Examples. :)
Param(
   [Parameter(Mandatory = $true, Position = 0)]
   [System.Xml.Linq.XName]$tag
,
   [Parameter(Position=99, Mandatory = $false, ValueFromRemainingArguments=$true)]
   [PSObject[]]$args
)
  Write-Verbose $($args | %{ $_ | Out-String } | Out-String)
  New-Object XElement $(
     $tag
     while($args) {
        $attrib, $value, $args = $args
        if($attrib -is [ScriptBlock]) {
           &$attrib
        } elseif ( $value -is [ScriptBlock] -and "-Content".StartsWith($attrib)) {
           &$value
        } elseif ( $value -is [XNamespace]) {
            New-XAttribute ([XNamespace]::Xmlns + $attrib.TrimStart("-")),$value
        } else {
           New-XAttribute $attrib.TrimStart("-"), $value
        }
     }
   )
}
Set-Alias xe New-XElement -Option AllScope
Set-Alias New-XmlElement New-XElement -Option AllScope

Export-ModuleMember -function New-XDocument, New-XAttribute, New-XElement, Remove-XmlNamespace, Convert-Xml, Select-Xml, Format-Xml -alias *