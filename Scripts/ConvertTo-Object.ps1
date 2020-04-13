##
##  this script converts hashtables to objects. Anything that's not a hashtable
##  is passed through the pipeline.
##

begin
{
	 Set-StrictMode -Off

	 if ($args[0] -eq '-?')
	 {
		  get-content $(get-command ConvertTo-Object.help.txt).Definition
		  exit
	 }
}


process
{
	 if ($_.GetType().FullName -ne "System.Collections.Hashtable")
	 {
		  return $_
	 }
	 $o = new-object PSObject
	 foreach ($key in $_.keys)
	 {
		  add-member -in $o -type NoteProperty -name $key -value $_[$key]
	 }
	 $o
}