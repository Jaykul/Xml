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
# Version    4.3 Added a Path parameter set to Format-XML so you can specify xml files for prety printing
# Version    4.5 Fixed possible [Array]::Reverse call on a non-array in New-XElement (used by New-XDocument)
#                Work around possible variable slipping on null values by:
#                1) allowing -param:$value syntax (which doesn't fail when $value is null)
#                2) testing for -name syntax on the value and using it as an attribute instead
# Version    4.6 Added -Arguments to Convert-Xml so that you can pass arguments to XSLT transforms!
#                Note: when using strings for xslt, make sure you single quote them or escape the $ signs.
# Version    4.7 Fixed a typo in the namespace parameter of Select-Xml
# Version    4.8 Fixed up some uses of Select-Xml -RemoveNamespace

$xlr8r = [psobject].assembly.gettype("System.Management.Automation.TypeAccelerators")
$xlinq = [Reflection.Assembly]::Load("System.Xml.Linq, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
$xlinq.GetTypes() | ? { $_.IsPublic -and !$_.IsSerializable -and $_.Name -ne "Extensions" -and !$xlr8r::Get[$_.Name] } | % {
  $xlr8r::Add( $_.Name, $_.FullName )
}
if(!$xlr8r::Get["Stack"]) {
   $xlr8r::Add( "Stack", "System.Collections.Generic.Stack``1, System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" )
}
if(!$xlr8r::Get["Dictionary"]) {
   $xlr8r::Add( "Dictionary", "System.Collections.Generic.Dictionary``2, mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" )
}
if(!$xlr8r::Get["PSParser"]) {
   $xlr8r::Add( "PSParser", "System.Management.Automation.PSParser, System.Management.Automation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" )
}



filter Format-XML {
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
Param(
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
}
Set-Alias fxml Format-Xml -EA 0

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
            $node.SelectNodes($path, $nsManager)
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
               $Xml = Remove-XmlNamespace -Xml $Xml
            }
            Select-Node $Xml $XPath $NSM
         }
      }
      "Xml" {
         foreach($node in $Xml) {
            if($RemoveNamespace) {
               [Xml]$node = Remove-XmlNamespace -Xml $node
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
         $Xml = Remove-XmlNamespace -Xml $Xml
      }
      Select-Node $Xml $XPath  $NSM
   }
}

}
Set-Alias slxml Select-Xml -EA 0

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
,
   [Alias("Parameters")]
   [hashtable]$Arguments   
)
BEGIN { 
   $StyleSheet = New-Object System.Xml.Xsl.XslCompiledTransform
   if(Test-Path @($Xslt)[0] -EA 0) { 
      Write-Verbose "Loading Stylesheet from $(Resolve-Path @($Xslt)[0])"
      $StyleSheet.Load( (Resolve-Path @($Xslt)[0]) )
   } else {
      $OFS = "`n"
      Write-Verbose "$Xslt"
      $StyleSheet.Load(([System.Xml.XmlReader]::Create((New-Object System.IO.StringReader "$Xslt" )) ))
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
            Convert-Node -Xml ([System.Xml.XmlReader]::Create((Resolve-Path $file))) $StyleSheet $Arguments
         }
      }
      "Xml" {
         foreach($node in $Xml) {
            Convert-Node -Xml (New-Object Xml.XmlNodeReader $node) $StyleSheet $Arguments
         }
      }
   }
}
END {
   if($PSCmdlet.ParameterSetName -eq "Content") {
      [Xml]$Xml = $XmlContent.ToString()
      Convert-Node -Node $Xml $StyleSheet $Arguments
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
#,
#   [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false)]
#   [ValidateNotNullOrEmpty()]
#   [Alias("StyleSheet")]
#   [String[]]$Xslt
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
            Convert-Node -Node $Xml $StyleSheet
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
      Convert-Node -Node $Xml $StyleSheet
   }
}
}
Set-Alias rmns Remove-XmlNamespace -EA 0



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
#  BEGIN {
   #  if([string]::IsNullOrEmpty( $tag.NamespaceName )) {
      #  $tag = $($script:NameSpaceStack.Peek()) + $tag
      #  if( $script:NameSpaceStack.Count -gt 0 ) {
         #  $script:NameSpaceStack.Push( $script:NameSpaceStack.Peek() )
      #  } else {
         #  $script:NameSpaceStack.Push( $null )
      #  }      
   #  } else {
      #  $script:NameSpaceStack.Push( $tag.Namespace )
   #  }
#  }
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
#  END {
   #  $null = $script:NameSpaceStack.Pop()
#  }
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


######## CliXml

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


######## Xaml
#  if($PSVersionTable.CLRVersion -ge "4.0"){
#     trap { continue }
#     [Reflection.Assembly]::LoadWithPartialName("System.Xaml") | Out-Null
#     if("System.Xaml.XamlServices" -as [type]) {
    
   #  }
#  }
   
Export-ModuleMember -alias * -function New-XDocument, New-XAttribute, New-XElement, Remove-XmlNamespace, Convert-Xml, Select-Xml, Format-Xml, ConvertTo-CliXml, ConvertFrom-CliXml

# SIG # Begin signature block
# MIIRDAYJKoZIhvcNAQcCoIIQ/TCCEPkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUlQ5L584Qz0SNfIkR1t7nIJfC
# Friggg5CMIIHBjCCBO6gAwIBAgIBFTANBgkqhkiG9w0BAQUFADB9MQswCQYDVQQG
# EwJJTDEWMBQGA1UEChMNU3RhcnRDb20gTHRkLjErMCkGA1UECxMiU2VjdXJlIERp
# Z2l0YWwgQ2VydGlmaWNhdGUgU2lnbmluZzEpMCcGA1UEAxMgU3RhcnRDb20gQ2Vy
# dGlmaWNhdGlvbiBBdXRob3JpdHkwHhcNMDcxMDI0MjIwMTQ1WhcNMTIxMDI0MjIw
# MTQ1WjCBjDELMAkGA1UEBhMCSUwxFjAUBgNVBAoTDVN0YXJ0Q29tIEx0ZC4xKzAp
# BgNVBAsTIlNlY3VyZSBEaWdpdGFsIENlcnRpZmljYXRlIFNpZ25pbmcxODA2BgNV
# BAMTL1N0YXJ0Q29tIENsYXNzIDIgUHJpbWFyeSBJbnRlcm1lZGlhdGUgT2JqZWN0
# IENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAyiOLIjUemqAbPJ1J
# 0D8MlzgWKbr4fYlbRVjvhHDtfhFN6RQxq0PjTQxRgWzwFQNKJCdU5ftKoM5N4YSj
# Id6ZNavcSa6/McVnhDAQm+8H3HWoD030NVOxbjgD/Ih3HaV3/z9159nnvyxQEckR
# ZfpJB2Kfk6aHqW3JnSvRe+XVZSufDVCe/vtxGSEwKCaNrsLc9pboUoYIC3oyzWoU
# TZ65+c0H4paR8c8eK/mC914mBo6N0dQ512/bkSdaeY9YaQpGtW/h/W/FkbQRT3sC
# pttLVlIjnkuY4r9+zvqhToPjxcfDYEf+XD8VGkAqle8Aa8hQ+M1qGdQjAye8OzbV
# uUOw7wIDAQABo4ICfzCCAnswDAYDVR0TBAUwAwEB/zALBgNVHQ8EBAMCAQYwHQYD
# VR0OBBYEFNBOD0CZbLhLGW87KLjg44gHNKq3MIGoBgNVHSMEgaAwgZ2AFE4L7xqk
# QFulF2mHMMo0aEPQQa7yoYGBpH8wfTELMAkGA1UEBhMCSUwxFjAUBgNVBAoTDVN0
# YXJ0Q29tIEx0ZC4xKzApBgNVBAsTIlNlY3VyZSBEaWdpdGFsIENlcnRpZmljYXRl
# IFNpZ25pbmcxKTAnBgNVBAMTIFN0YXJ0Q29tIENlcnRpZmljYXRpb24gQXV0aG9y
# aXR5ggEBMAkGA1UdEgQCMAAwPQYIKwYBBQUHAQEEMTAvMC0GCCsGAQUFBzAChiFo
# dHRwOi8vd3d3LnN0YXJ0c3NsLmNvbS9zZnNjYS5jcnQwYAYDVR0fBFkwVzAsoCqg
# KIYmaHR0cDovL2NlcnQuc3RhcnRjb20ub3JnL3Nmc2NhLWNybC5jcmwwJ6AloCOG
# IWh0dHA6Ly9jcmwuc3RhcnRzc2wuY29tL3Nmc2NhLmNybDCBggYDVR0gBHsweTB3
# BgsrBgEEAYG1NwEBBTBoMC8GCCsGAQUFBwIBFiNodHRwOi8vY2VydC5zdGFydGNv
# bS5vcmcvcG9saWN5LnBkZjA1BggrBgEFBQcCARYpaHR0cDovL2NlcnQuc3RhcnRj
# b20ub3JnL2ludGVybWVkaWF0ZS5wZGYwEQYJYIZIAYb4QgEBBAQDAgABMFAGCWCG
# SAGG+EIBDQRDFkFTdGFydENvbSBDbGFzcyAyIFByaW1hcnkgSW50ZXJtZWRpYXRl
# IE9iamVjdCBTaWduaW5nIENlcnRpZmljYXRlczANBgkqhkiG9w0BAQUFAAOCAgEA
# UKLQmPRwQHAAtm7slo01fXugNxp/gTJY3+aIhhs8Gog+IwIsT75Q1kLsnnfUQfbF
# pl/UrlB02FQSOZ+4Dn2S9l7ewXQhIXwtuwKiQg3NdD9tuA8Ohu3eY1cPl7eOaY4Q
# qvqSj8+Ol7f0Zp6qTGiRZxCv/aNPIbp0v3rD9GdhGtPvKLRS0CqKgsH2nweovk4h
# fXjRQjp5N5PnfBW1X2DCSTqmjweWhlleQ2KDg93W61Tw6M6yGJAGG3GnzbwadF9B
# UW88WcRsnOWHIu1473bNKBnf1OKxxAQ1/3WwJGZWJ5UxhCpA+wr+l+NbHP5x5XZ5
# 8xhhxu7WQ7rwIDj8d/lGU9A6EaeXv3NwwcbIo/aou5v9y94+leAYqr8bbBNAFTX1
# pTxQJylfsKrkB8EOIx+Zrlwa0WE32AgxaKhWAGho/Ph7d6UXUSn5bw2+usvhdkW4
# npUoxAk3RhT3+nupi1fic4NG7iQG84PZ2bbS5YxOmaIIsIAxclf25FwssWjieMwV
# 0k91nlzUFB1HQMuE6TurAakS7tnIKTJ+ZWJBDduUbcD1094X38OvMO/++H5S45Ki
# 3r/13YTm0AWGOvMFkEAF8LbuEyecKTaJMTiNRfBGMgnqGBfqiOnzxxRVNOw2hSQp
# 0B+C9Ij/q375z3iAIYCbKUd/5SSELcmlLl+BuNknXE0wggc0MIIGHKADAgECAgFR
# MA0GCSqGSIb3DQEBBQUAMIGMMQswCQYDVQQGEwJJTDEWMBQGA1UEChMNU3RhcnRD
# b20gTHRkLjErMCkGA1UECxMiU2VjdXJlIERpZ2l0YWwgQ2VydGlmaWNhdGUgU2ln
# bmluZzE4MDYGA1UEAxMvU3RhcnRDb20gQ2xhc3MgMiBQcmltYXJ5IEludGVybWVk
# aWF0ZSBPYmplY3QgQ0EwHhcNMDkxMTExMDAwMDAxWhcNMTExMTExMDYyODQzWjCB
# qDELMAkGA1UEBhMCVVMxETAPBgNVBAgTCE5ldyBZb3JrMRcwFQYDVQQHEw5XZXN0
# IEhlbnJpZXR0YTEtMCsGA1UECxMkU3RhcnRDb20gVmVyaWZpZWQgQ2VydGlmaWNh
# dGUgTWVtYmVyMRUwEwYDVQQDEwxKb2VsIEJlbm5ldHQxJzAlBgkqhkiG9w0BCQEW
# GEpheWt1bEBIdWRkbGVkTWFzc2VzLm9yZzCCASIwDQYJKoZIhvcNAQEBBQADggEP
# ADCCAQoCggEBAMfjItJjMWVaQTECvnV/swHQP0FTYUvRizKzUubGNDNaj7v2dAWC
# rAA+XE0lt9JBNFtCCcweDzphbWU/AAY0sEPuKobV5UGOLJvW/DcHAWdNB/wRrrUD
# dpcsapQ0IxxKqpRTrbu5UGt442+6hJReGTnHzQbX8FoGMjt7sLrHc3a4wTH3nMc0
# U/TznE13azfdtPOfrGzhyBFJw2H1g5Ag2cmWkwsQrOBU+kFbD4UjxIyus/Z9UQT2
# R7bI2R4L/vWM3UiNj4M8LIuN6UaIrh5SA8q/UvDumvMzjkxGHNpPZsAPaOS+RNmU
# Go6X83jijjbL39PJtMX+doCjS/lnclws5lUCAwEAAaOCA4EwggN9MAkGA1UdEwQC
# MAAwDgYDVR0PAQH/BAQDAgeAMDoGA1UdJQEB/wQwMC4GCCsGAQUFBwMDBgorBgEE
# AYI3AgEVBgorBgEEAYI3AgEWBgorBgEEAYI3CgMNMB0GA1UdDgQWBBR5tWPGCLNQ
# yCXI5fY5ViayKj6xATCBqAYDVR0jBIGgMIGdgBTQTg9AmWy4SxlvOyi44OOIBzSq
# t6GBgaR/MH0xCzAJBgNVBAYTAklMMRYwFAYDVQQKEw1TdGFydENvbSBMdGQuMSsw
# KQYDVQQLEyJTZWN1cmUgRGlnaXRhbCBDZXJ0aWZpY2F0ZSBTaWduaW5nMSkwJwYD
# VQQDEyBTdGFydENvbSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eYIBFTCCAUIGA1Ud
# IASCATkwggE1MIIBMQYLKwYBBAGBtTcBAgEwggEgMC4GCCsGAQUFBwIBFiJodHRw
# Oi8vd3d3LnN0YXJ0c3NsLmNvbS9wb2xpY3kucGRmMDQGCCsGAQUFBwIBFihodHRw
# Oi8vd3d3LnN0YXJ0c3NsLmNvbS9pbnRlcm1lZGlhdGUucGRmMIG3BggrBgEFBQcC
# AjCBqjAUFg1TdGFydENvbSBMdGQuMAMCAQEagZFMaW1pdGVkIExpYWJpbGl0eSwg
# c2VlIHNlY3Rpb24gKkxlZ2FsIExpbWl0YXRpb25zKiBvZiB0aGUgU3RhcnRDb20g
# Q2VydGlmaWNhdGlvbiBBdXRob3JpdHkgUG9saWN5IGF2YWlsYWJsZSBhdCBodHRw
# Oi8vd3d3LnN0YXJ0c3NsLmNvbS9wb2xpY3kucGRmMGMGA1UdHwRcMFowK6ApoCeG
# JWh0dHA6Ly93d3cuc3RhcnRzc2wuY29tL2NydGMyLWNybC5jcmwwK6ApoCeGJWh0
# dHA6Ly9jcmwuc3RhcnRzc2wuY29tL2NydGMyLWNybC5jcmwwgYkGCCsGAQUFBwEB
# BH0wezA3BggrBgEFBQcwAYYraHR0cDovL29jc3Auc3RhcnRzc2wuY29tL3N1Yi9j
# bGFzczIvY29kZS9jYTBABggrBgEFBQcwAoY0aHR0cDovL3d3dy5zdGFydHNzbC5j
# b20vY2VydHMvc3ViLmNsYXNzMi5jb2RlLmNhLmNydDAjBgNVHRIEHDAahhhodHRw
# Oi8vd3d3LnN0YXJ0c3NsLmNvbS8wDQYJKoZIhvcNAQEFBQADggEBACY+J88ZYr5A
# 6lYz/L4OGILS7b6VQQYn2w9Wl0OEQEwlTq3bMYinNoExqCxXhFCHOi58X6r8wdHb
# E6mU8h40vNYBI9KpvLjAn6Dy1nQEwfvAfYAL8WMwyZykPYIS/y2Dq3SB2XvzFy27
# zpIdla8qIShuNlX22FQL6/FKBriy96jcdGEYF9rbsuWku04NqSLjNM47wCAzLs/n
# FXpdcBL1R6QEK4MRhcEL9Ho4hGbVvmJES64IY+P3xlV2vlEJkk3etB/FpNDOQf8j
# RTXrrBUYFvOCv20uHsRpc3kFduXt3HRV2QnAlRpG26YpZN4xvgqSGXUeqRceef7D
# dm4iTdHK5tIxggI0MIICMAIBATCBkjCBjDELMAkGA1UEBhMCSUwxFjAUBgNVBAoT
# DVN0YXJ0Q29tIEx0ZC4xKzApBgNVBAsTIlNlY3VyZSBEaWdpdGFsIENlcnRpZmlj
# YXRlIFNpZ25pbmcxODA2BgNVBAMTL1N0YXJ0Q29tIENsYXNzIDIgUHJpbWFyeSBJ
# bnRlcm1lZGlhdGUgT2JqZWN0IENBAgFRMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEWMCMGCSqGSIb3DQEJBDEWBBTM4lHUVRUC
# Gv/aYFU0KwQT41++izANBgkqhkiG9w0BAQEFAASCAQAhor+qcezuBnhODNs9y4vb
# 3gP5gdOUAbPJWAaV82VNO0pQcqahwT/PyXl9UZIacmy5hQ8UVAIKK/+Y/r6Wcz2+
# Pe/kh0cvD8n9PFU1UIHHj8ucyS7lzHrLzojulx2zy3Wtm8Lo4t8IrY8jNl1sSH+O
# 8XP1Rgv+3CkIoERKkUpcPEWruj5kJyyHE+sRNQcZGlO9gjx8dzoUlLyxQXIF+fmW
# ix2qQJ4UC6B1pjjv/Fv5KMuSDpgLOhLtzQeubmeqBGjeoumaPvYgvPsn6rYvBxWY
# 9wdP8VCGVCIrEsUUyAuXoxzpFJk80PAX08bdvcOkQ7GQ13HEN4JBDtDTYHFV9M7v
# SIG # End signature block