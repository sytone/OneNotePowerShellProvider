##
##  This script imports files into a OneNote page or section.
##

param(
     $OneNotePath,
     $FileSpec,
     [switch] $embed,
     $SubstituteContentExtension
     )

if ($args[0] -eq "-?")
{
	 Get-Content $(Get-Command Import-FilesToOneNote.help.txt).Definition
	 return
}

if (!$OneNotePath)
{
	 throw "You must enter the path to a OneNote section."
}

if (!$FileSpec)
{
	 throw "You must enter one or more files to import."
}


##
##  Helper routine that does the actual work in the pipeline.
## 

filter Add-FileToPage
{
	 $fileNameWithoutExtension = $_.Name.Replace( $_.Extension, "" )
	 write-verbose $fileNameWithoutExtension
	 $pageName = "$OneNotePath\$fileNameWithoutExtension"
	 write-verbose $pageName
	 new-item $pageName -type page
	 if ($embed)
	 {
		  "############################################################" | add-content $pageName
		  $_ | add-content $pageName
		  "############################################################" | add-content $pageName
	 }
	 $substitute = $_.FullName.Replace( $_.Extension, $SubstituteContentExtension )
	 write-verbose "Looking for $substitute"
	 if ($(test-path $substitute))
	 {
		  get-content $substitute | add-content $pageName
	 } else
	 {
		  get-content $_ | add-content $pageName
	 }
}

Get-Item $FileSpec | Add-FileToPage



