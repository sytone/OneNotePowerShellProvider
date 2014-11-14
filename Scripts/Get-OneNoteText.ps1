##
##  Gets just the text off of a OneNote page.
##

param( $OneNotePath, $Stylesheet )

if ($args[0] -eq "-?")
{
	 Get-Content $(Get-Command Get-OneNoteText.help.txt).Definition
	 return
}

if (!$OneNotePath)
{
	 throw "You must enter the path to a OneNote page or a file in the file system."
}

##
##  Hack -- if $OneNotePath isn't a OneNote file, then just Get-Content the path.
##

$i = Get-Item $OneNotePath
if ($i.PSObject.TypeNames[0] -ne 'System.Xml.XmlElement#http://schemas.microsoft.com/office/onenote/2007/onenote#Page')
{
	 Get-Content $OneNotePath
	 return
}

##
##  Test for prerequisites (from PowerShell Community Extensions)
##

function test-command ($commandName)
{
	 $error = $($commands = get-command $commandName) 2>&1
	 return $commands
}

if (!$(test-command Convert-Xml))
{
	 throw "The PowerShell Community Extensions do not appear to be installed. `nYou must install them to use this script."
}

if (!$Stylesheet)
{
	 $Stylesheet = join-path $(split-path $MyInvocation.MyCommand.Definition -parent) `
	     "OnToText.xslt"
	 write-verbose $Stylesheet
}

$text = Get-OneNotePageContent $OneNotePath | Convert-Xml -xslt $Stylesheet

##
##  Remove HTML markup.
##

$markup = [regex] "<[^>]*>"
$text = $markup.Replace( $text, "" )

##
##  common entity conversions
##

$text = $text.Replace( "&amp;", "&" )

##
##  Emit the text
##

$text.Trim() | split-string -newline


