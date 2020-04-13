##
##  Starts a series of tests. An individual test is an object with the following properties:
##    Name -- the name of the test
##    Description -- the description of the test
##    ScriptBlock -- what you execute to do the test
##    ErrorAction -- OPTIONAL, what to do in the event of an error
##    ExpectedString -- OPTIONAL, the expected output from executing the test as a string
##
param( $Filter,
       $encoding = 'ascii',
       [switch] $OnlyFilteredTests,
       [switch] $Transcript )

begin
{

	 if ($args[0] -eq "-?")
	 {
		  Get-Content $(Get-Command Start-Tests.help.txt).Definition
		  exit
	 }

	 ##
	 ##  Functions for managing the transcript
	 ##

	 function prepend-CommentCharacters()
	 {
		  begin {}
		  process 
		  {
			   return "##  " + $_
		  }
		  end {}
	 }
	 
	 function write-TranscriptComment( $strings )
	 {
		  if ($Transcript)
		  {
			   (" ", $strings, " ") | prepend-CommentCharacters | 
		            out-file -filePath $TranscriptFile -append -encoding $encoding
			   out-file -inputObject "`n" -filePath $TranscriptFile -append -encoding $encoding
		  }
	 }
	 
	 function write-TranscriptCommand( $strings )
	 {
		  if ($Transcript)
		  {
			   (" ", $strings, " ") |  
		       out-file -filePath $TranscriptFile -append -encoding $encoding
			   out-file -inputObject "`n" -filePath $TranscriptFile -append -encoding $encoding
		  }
	 }
	 
	 function trim-string()
	 {
		  begin {}
		  process { return $_.trim() }
		  end {}
	 }
	 
	 function test-filter( $filterTerm )
	 {
		  $results = @()
		  foreach ($s in $Filter)
		  {
			   if ($filterTerm -ilike "$s*")
			   {
					$results += $s
			   }
		  }
		  return $results
	 }
	 
	 ##
	 ##  Initialize the transcript
	 ##
	 
	 if ($Transcript)
	 {
		  $TranscriptFile = join-path $(get-location) "transcript.ps1"
		  write-host "Recording transcript in $TranscriptFile." -fore yellow
		  new-item $TranscriptFile -type file -force
	 }
}

process
{
	 if ($_.Filter)
	 {
		  if (!(test-filter( $_.Filter )))
		  {
			   write-host ("Skipping test ", $_.Name, " because of filter ", $_.Filter) `
			        -fore yellow -back black
			   return
		  }
	 }
	 if ($OnlyFilteredTests -and !$_.Filter)
	 {
		  write-host "Skipping test " $_.name " because it is not filtered." -sep $null `
		      -fore yellow -back black
		  return
	 }
	 write-host ("Starting test: ", $_.Name) -fore Yellow
	 write-host 
	 write-host $_.ScriptBlock.ToString().Trim()

	 $invokeError = $($results = & $_.ScriptBlock) 2>&1
	 if ($_.ValidateError)
	 {
		  ##
		  ##  the ValidateError scriptblock gives the test author an opportunity
		  ##  to gobble up any errors that may have happened when executing the last
		  ##  test case. You can also guarantee that errors happen when you expect them.
		  ##

		  $invokeError = $($invokeError | & $_.ValidateError) 2>&1
	 }
	 if ($invokeError)
	 {
		  write-host "Error on invoke:", $invokeError -fore Red -back Black
		  if ($_.ErrorAction -inotlike "continue")
		  {
			   break
		  }
	 }
	 $resultString = $results | out-string
	 $resultString = $resultString.Trim()
	 write-host "`n", $resultString, "`n" -sep $null
	 if ($_.ExpectedString)
	 {
		  $expected = $_.ExpectedString
		  $expected = $expected.Trim()
		  if ($resultString -ne $expected)
		  {
			   write-host ("Expected results:`n", $expected) -fore Red -sep $null
			   if ($_.ErrorAction -inotlike "continue")
			   {
					break
			   }
		  }
	 }
	 if ($_.Validate)
	 {
		  $validateError = $($results | & $_.Validate) 2>&1
		  if ($validateError)
		  {
			   write-host "Validation error: $validateError" -fore Red 
			   if ($_.ErrorAction -inotlike "continue")
			   {
					break
			   }
		  }
	 }

	 ##
	 ##  Maintain the transcript
	 ## 

	 if ($Transcript)
	 {
		  (" ", $_.Name, " ", $_.Description, " ") | split-string -newline | trim-string |
		      prepend-CommentCharacters | out-file -file $TranscriptFile -append -encoding $encoding
		  (" ", $_.ScriptBlock, " ") | split-string -newline | trim-string | out-file -file $TranscriptFile -append -encoding $encoding
		  ("##", "##  Results:") | out-file -file $TranscriptFile -append -encoding $encoding
		  $results | out-string | split-string -newline | prepend-CommentCharacters | out-file -file $TranscriptFile -append -encoding $encoding
		  (" ", " ") | out-file -file $TranscriptFile -append -encoding $encoding
	 }
}

end
{
}

