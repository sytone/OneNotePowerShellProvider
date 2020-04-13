##
##  This script is the PowerShell equivalent of JSON notation -- it exports objects
##  in the pipeline to PowerShell syntax that will let them get reconstituted by
##  executing the particular script.
##

param( $path, $encoding = 'ASCII', [switch]$passThru, [switch]$stdout )



begin
{
	 Set-StrictMode -Off

	 if ($args[0] -eq "-?")
	 {
		  Get-Content $(Get-Command Export-PsOn.help.txt).Definition
		  exit
	 }
	 function ConvertTo-PsON( $inputObject, $indent )
	 {
		  switch ($inputObject.GetType().FullName)
		  {
			   "System.String"
			   {
					$escaped = $inputObject.Replace( "'", "''" )
					"'$escaped'"
			   }
			   "System.DateTime"
			   {
					"[datetime] '$($inputObject.ToString('s'))'"
			   }
			   "System.Boolean"
			   {
					if ($inputObject)
					{
						 '$true'
					} else
					{
						 '$false'
					}
			   }
			   "System.Management.Automation.ScriptBlock"
			   {
					"{`n"
					"$indent$x"
					"}"
			   }
			   "System.Int32"
			   {
					"[int] '$($inputObject.toString())'"
			   }
			   "System.Double"
			   {
					"[double] '$($inputObject.toString())'"
			   }
			   "System.Collections.Hashtable"
			   {
					"@{`n"
					foreach ($key in $inputObject.keys)
					{
						 $escaped = $key.Replace( "'", "''" )
						 "$indent'$escaped'="
						 ConvertTo-PsON -inputObject $inputObject[$key] -indent "$indent  "
						 ";`n"
					}
					"}"
			   }
			   "System.Object[]"
			   {
					for ($i = 0; $i -lt $inputObject.length; $i++)
					{
						 ConvertTo-PsON -inputObject $inputObject[$i] -indent "$indent  "
						 if ($i -ne ($inputObject.length - 1))
						 {
							  ', '
						 }
					}
			   }
			   {($_ -eq "System.Xml.XmlElement") -or ($_ -eq "System.Xml.XmlDocument") -or ($_ -eq "System.Management.Automation.PSCustomObject")}
			   {
					"@{`n"
					foreach ($property in $($inputObject | get-member -membertype property, NoteProperty))
					{
						 "$indent$($property.name)="
						 ConvertTo-PsOn -inputObject $inputObject.$($property.name) -indent "$indent  "
						 ";`n"
					}
					"}"
			   }
			   default
			   {
					throw "Unrecognized type: $($inputObject.GetType().FullName)"
			   }
		  }
	 }
	 $output = new-object System.Text.StringBuilder
}

process
{
	 $strings = ConvertTo-PsOn -inputObject $_ -indent "  "
	 foreach ($string in $strings)
	 {
		  $output.Append( "$string" ) | out-null
	 }
	 $output.Append( "`n" ) | out-null
	 if ($passThru)
	 {
		  $_
	 }
}

end
{
	 $str = $output.ToString()
	 if ($path)
	 {
		  $str | out-file $path -encoding $encoding
	 }
	 if ($stdout)
	 {
		  out-host -in $str
	 }
}
