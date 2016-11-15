# XML - A Module by Joel Bennett

A module providing converters for HTML to XML, various core XML commands and most importantly, _a DSL for generating XML documents_.

Two commands to convert HTML into XML
-------------------------------------

#### Import-Html
> Load an HTML file to an XML Document

#### ConvertFrom-Html
> Load HTML text as an XML Document


Commands for working with XML
-----------------------------

#### Set-XmlContent
> Save an XmlDocument or Node to the specified file path

#### Get-XmlContent
> Load an XML file as an XmlDocument

#### Convert-Xml
> The Convert-XML function lets you use Xslt to transform XML strings and documents

#### Format-Xml
> Pretty-print formatted XML source

#### Get-XmlContent
> Load an XML file as an XmlDocument

#### Set-XmlContent
> Save an XmlDocument or Node to the specified file path

#### Remove-XmlElement
> Removes specified elements (tags or attributes) or all elements from a specified namespace from an Xml document

#### Remove-XmlNamespace
> Removes namespace definitions and prefixes from xml documents

#### Select-Xml
> The Select-XML cmdlet lets you use XPath queries to search for text in XML strings and documents. Enter an XPath query, and use the Content, Path, or Xml parameter to specify the XML to be searched

#### Update-Xml
> The Update-XML cmdlet lets you use XPath queries to replace text in nodes in XML documents. Enter an XPath query, and use the Content, Path, or Xml parameter to specify the XML to be searched


Helpers for working with CliXml without files
---------------------------------------------

#### ConvertFrom-CliXml
> Imports CliXml content and creates corresponding objects within Windows PowerShell.

#### ConvertTo-CliXml
> Creates an CliXml-based representation of an object or objects and outputs it


The XML DSL
===========

Simply put, write something that looks like PowerShell, but each command is converted to an XML element, and each parameter/value is turned into an attribute and it's value.

It really has only one built-in command: `New-XDocument` (better known by it's alias: `XML`), which creates the new XML document!

Any "command" _that's not a real command_ turns into an XML node, and any parameters turn into attributes. 

Passing a ScriptBlock to these fake commands creates nested elements.

At any point (even nested several levels deep in the XML document) you can call PowerShell commands to output objects and convert them into XML nested inside your document.

XML DSL Commands
-----------------------------

#### New-XDocument (Alias: `XML`)
> Creates a new XML document [System.Xml.Linq.XDocument] from XmlDsl or a standard ScriptBlock.

#### New-XElement
> Creates a new XML element [System.Xml.Linq.XElement] from XmlDsl or a standard ScriptBlock.

#### ConvertFrom-XmlDsl
> Converts XmlDsl to a standard ScriptBlock.  If you are creating individual documents, ignore this command and send your XmlDsl directly to the New-X* commands above.  However, for improved performance when calling New-X* commands multiple times with the same XmlDsl template, use the output of this function for those New-X* commands combined with -BlockType Script to bypass the overhead of conversion.  

Version History:
================

* **2.0** This was the first script version I wrote. It didn't function identically to the built-in Select-Xml with regards to parameter parsing
* **2.1** Matched the built-in Select-Xml parameter sets, it's now a drop-in replacement, BUT only if you were using the original with: Select-Xml ... | Select-Object -Expand Node
* **2.2** Fixes a bug in the -Content parameterset where -RemoveNamespace was *presumed* 
* **3.0** Added New-XDocument and associated generation functions for my XML DSL
* **3.1** Fixed a really ugly bug in New-XDocument in 3.0 which I should not have released
* **4.0** Never content to leave well enough alone, I completely reworked New-XDocument
* **4.1** Tweaked namespaces again so they don't cascade down when they shouldn't. Got rid of the unnecessary stack.
* **4.2** Tightened xml: only cmdlet, function, and external scripts, with "-" in their names are exempted from being converted into xml tags. Fixed some alias error messages caused when PSCX is already loaded (we overwrite their aliases for cvxml and fxml)
* **4.3** Added a Path parameter set to Format-Xml so you can specify xml files for prety printing
* **4.5** Fixed possible [Array]::Reverse call on a non-array in New-XElement (used by New-XDocument). Worked around possible variable slipping on null values by:
  1. Allowing -param:$value syntax (which doesn't fail when $value is null)
  2. testing for -name syntax on the value and using it as an attribute instead
* **4.6** Added -Arguments to Convert-Xml so that you can pass arguments to XSLT transforms! Note: when using strings for xslt, make sure you single quote them or escape the $ signs.
* **4.7** Fixed a typo in the namespace parameter of Select-Xml
* **4.8** Fixed up some uses of Select-Xml -RemoveNamespace
* **5.0** Added Update-Xml to allow setting xml attributes or node content
* **6.0** Major cleanup, breaking changes.
  * Added Get-XmlContent and Set-XmlContent for loading/saving XML from files or strings
  * Removed Path and Content parameters from the other functions (it greatly simplifies those functions, and makes the whole thing more maintainable)
  * Updated Update-Xml to support adding nodes "before" and "after" other nodes, and to support "remove"ing nodes
* **6.1** Update for PowerShell 3.0
* **6.2** Minor tweak in exception handling for CliXml
* **6.3** Added Remove-XmlElement to allow removing nodes or attributes. This is something I specifically needed to remove "ignorable" namespaces such as the ones created by the Visual Studio Workflow designer (and perhaps other visual designers like Blend) which I don't want to check into source control, because it makes diffing nearly impossible.
* **6.4** Fixed a bug on New-XElement for Rudy Shockaert (nice bug report, thanks!)
* **6.5** Added -Parameters @{} parameter to New-XDocument to allow local variables to be passed into the module scope. *grumble*
* **6.6** Expose Convert-Xml and fix a couple of bugs (I can't figure how they got here)
* **6.7** Add ConvertFrom-Html, add -Formatted switch to Set-XmlContent. Reformat and clean up to (mostly) match the new best practices I'm working on. Finally put the module on github.
* **7.0** Update to publish to PowerShell Gallery