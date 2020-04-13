##
##  This script "starts" the OneNote provider. You should call it every time you start
##  powershell. It's a great candidate for including in your profile.
##

param([switch]$silent)

if ($args[0] -eq "-?")
{
	 Get-Content $(Get-Command Enable-OneNote.help.txt).Definition
	 return
}

function test-command ($commandName)
{
	 $error = $($commands = get-command $commandName) 2>&1
	 return $commands
}

$OneNoteHome = $(split-path $MyInvocation.MyCommand.Definition -parent)
write-verbose $OneNoteHome
if (!$silent -and $(test-command "Get-FileVersionInfo"))
{
	 write-verbose "Getting file version info for $OneNoteHome\*.dll"
	 Get-FileVersionInfo -Path "$OneNoteHome\*.dll" | format-table ProductName, FileVersion
}


Import-Module "$($OneNoteHome)\Microsoft.Office.OneNote.PowerShell.dll"
Update-FormatData "$($OneNoteHome)\OneNote.ps1xml"
$global:OneNoteHome = $OneNoteHome

