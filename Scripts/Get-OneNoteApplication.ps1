##
##  Gets a handle to a OneNote application object
##

resolve-assembly "Microsoft.Office.Interop.OneNote" -import | out-null
$global:app = New-Object -type "Microsoft.Office.Interop.OneNote.ApplicationClass"
return $global:app

