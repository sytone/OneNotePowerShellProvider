##
##  Gets a digest of the changed pages for a OneNote notebook. Can optionally
##  send these pages in email.
##

param( $Notebook,
    $Container = ${env:temp},
    [datetime]$targetDate = $(get-date).AddDays(-1),
    $Stylesheet,
    $TextStylesheet,
    $MailTo = "${env:username}@microsoft.com",
    $MailFrom = "${env:username}@microsoft.com",
    [switch]$whatIf,
    [switch]$noClean,
    [switch]$Verbose )

if ($args[0] -eq '-?')
{
	 Get-Content $(Get-Command Get-OneNoteDigest.help.txt).Definition
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

if (!$(test-command Convert-Xml) -or !$(test-command Send-SMTPMail))
{
	 throw "The PowerShell Community Extensions do not appear to be installed. `nYou must install them to use this script."
}

if (!$Stylesheet)
{
	 $Stylesheet = join-path $(split-path $MyInvocation.MyCommand.Definition -parent) `
	     "OfficeLabsDigest.xslt"
	 write-verbose $Stylesheet
}

if (!$TextStylesheet)
{
	 $TextStylesheet = join-path $(split-path $MyInvocation.MyCommand.Definition -parent) `
	     "OnToText.xslt"
	 write-verbose $TextStylesheet
}

if (!$Notebook)
{
	 throw "You must enter the path to a OneNote item."
}

if ($Verbose)
{
	 $savedPref = $global:VerbosePreference
	 $global:VerbosePreference = "Continue"
}

##
##  Helper function
##

function escape-string()
{
	 begin{}
	 process 
	 {
		  $_ = $_.Replace( "&", "&amp;" )
		  $_ = $_.Replace( "<", "&lt;" )
		  $_ = $_.Replace( ">", "&gt;" )
		  return $_
	 }
	 end {}
}

##
##  This function filters out pages that haven't *really* changed by
##  comparing the hash of the page content (text only) with the last
##  hash seen. This is to work through a bug in OneNote where some pages
##  ALWAYS show up as changed.
##

function Remove-UnchangedPages( $hashFile, $stylesheet )
{
	 begin
	 {
		  if (test-path $hashFile)
		  {
			   $hashes = & $hashFile
		  } else
		  {
			   $hashes = @{ }
		  }
	 }
	 process
	 {
		  $hash = Get-Hash -InputObject $(Get-OneNotePageContent -id $_.id | Convert-Xml -xslt $stylesheet)
		  if ($hashes[ $_.pspath] -and ($hashes[ $_.pspath ] -eq $hash.HashString))
		  {
			   write-verbose "Page $($_.pspath) has not really changed. Skipping."
			   return
		  }
		  $hashes[ $_.pspath ] = $hash.HashString
		  return $_
	 }
	 end
	 {
		  $hashes | Export-PsOn -path $hashFile
	 }
}

##
##  I need a working directory. It will get created in $Container, which defaults to
##  %temp%.
##

$notebookName = $(Get-OneNoteHierarchy $Notebook -scope hsSelf).name
$outputDirectory = join-path $Container $notebookName
new-item $outputDirectory -type directory -force | out-null

##
##  Now, get all of the pages that have changed since $targetDate.
##

$changedPages = dir $Notebook -recurse |
    where-object { ([datetime]$_.lastModifiedTime -gt $targetDate) -and (!$_.PSIsContainer) } |
    Remove-UnchangedPages -hashFile "$outputDirectory\PageHashes.ps1" -style $TextStylesheet

if ($changedPages -eq $null)
{
	 if ($Verbose) { $global:VerbosePreference = $savedPref }
     return "0 changed pages. No mail sent."
}

##
##  Force us into an array
##

$changedPages = @($changedPages)

write-verbose ($changedPages.length.toString( ) + " pages have changed since " + `
     $targetDate.ToShortDateString( ))

$outputDirectory = join-path $outputDirectory $(get-date).ToString( "yyyy-MM-ddTHH.mm.ss" )
new-item $outputDirectory -type directory -force | out-null
write-verbose ("Output will go to " + $outputDirectory)


##
##  Step 1: Convert the pages to MHT
##

$exportedFileNames = $changedPages | export-onenote -output $outputDirectory -format mht | 
    get-propertyvalue ExportedFile
$exportedFileNames | out-string | write-verbose

##
##  Step 2: Get the hyperlinks and store them in an array.
##

$hyperlinks = ($changedPages | Get-OneNoteHyperlink | escape-string)
$hyperlinks | out-string | write-verbose

##
##  Step 3: Write out an XML manifest of the changed pages and their hyperlinks.
##

$manifest = "<OneNoteDigest><NotebookName>$notebookName</NotebookName>"
$manifest += "<ChangesSince>" + $targetDate.ToShortDateString( ) + "</ChangesSince>"

for ($i = 0; $i -lt $changedPages.length; $i++)
{
	 $manifest += "<ChangedPage><Name><![CDATA[$($changedPages[$i].name)]]></Name>"
	 $manifest += "<Hyperlink>$($hyperlinks[$i])</Hyperlink>"
	 $manifest += "</ChangedPage>"
}
$manifest += "</OneNoteDigest>"

write-debug $manifest
($doc = [xml] "<root />").LoadXml( $manifest )
$doc.Save( "$outputDirectory\manifest.xml" )

##
##  Transform the XML to get the HTML body of the message.
##

Convert-Xml -Path "$outputDirectory\manifest.xml" -XsltPath $Stylesheet > `
    "$outputDirectory\manifest.htm"

##
##  Send the mail message.
##

$body = gc "$outputDirectory\manifest.htm" -Raw
if (!$whatIf)
{
	 Send-SmtpMail -to $MailTo -from bdewey@exchange.microsoft.com -subject "$notebookName Digest" `
          -body $body -html -attachmentLiteralPath $exportedFileNames -SmtpHost smtphost
}
"$($changedPages.length) changed page(s). Mail sent to $MailTo."
if (!$noClean)
{
	 sleep 3
	 $eatErrors = $(remove-item $outputDirectory -recurse -force -exclude *.ps1) 2>&1
}

##
##  Restore global verbose preferences
##
if ($Verbose)
{
	 $global:VerbosePreference = $savedPref
}


