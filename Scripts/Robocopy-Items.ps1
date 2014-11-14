##
##  This script invokes Robocopy to copy items from one location to another.
##  It is used as part of the web notebook publishing process.
##

process
{
	 if (($_.ExportedFile) -and ($_.DestinationPath))
	 {
		  Robocopy $_.ExportedFile $_.DestinationPath /s
	 }
}
