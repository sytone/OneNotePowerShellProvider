##
##  The project is now quasi-self-documenting. This script will create a OneNote
##  notebook that contains the relevant documentation for the project. The notebook 
##  is then exported to a ONEPKG file.
##

if ($args[0] -eq '-?')
{
	 Get-Content $(Get-Command Create-OneNoteDocumentation.help.txt).Definition
	 return
}

$notebook = 'OneNote:\OneNote PowerShell Documentation'
new-item $notebook -type notebook -value ${env:TEMP}

##
##  Create the "About" section.
##

new-item "$notebook\About" -type section
new-item "$notebook\About\About the OneNote PowerShell Provider" -type page

Get-Help about_OneNote | set-content "$notebook\About\About the OneNote PowerShell Provider"
Get-ProviderTests | start-tests -transcript
new-item "$notebook\About\Demo" -type page
Get-Content transcript.ps1 | set-content "$notebook\About\Demo"

##
##  Create the "Scripts" section.
##

new-item "$notebook\Scripts" -type Section
Import-FilesToOneNote -OneNote "$notebook\Scripts" -File "$OneNoteHome\Scripts\*.ps1" `
    -Substitute ".help.txt" -embed

##
##  Create the "Cmdlets" section.
##

new-item "$notebook\Cmdlets" -type Section

filter Add-CmdletDocumentation
{
	 write-verbose "Adding documentation for cmdlet $($_.name)"
	 new-item "$notebook\Cmdlets\$($_.name)" -type page
	 Get-Help $_.name -detailed | out-string | add-content "$notebook\Cmdlets\$($_.name)"
}

Get-Command -PSSnapin Microsoft.Office.OneNote | Add-CmdletDocumentation

$results = Export-OneNote $notebook -format OnePkg -output "${env:temp}\OneNote Powershell Documentation.onepkg"

"Output is in $($results.ExportedFile)"
Close-OneNote $notebook
remove-item "${env:temp}\OneNote PowerShell Documentation" -recurse -force

