##
##  Turns OneNote debugging on or off.
##

param( [bool] $EnableDebugging = $true )

$key = 'hkcu:\software\microsoft\office\12.0\onenote\options\logging'
if ($EnableDebugging)
{
	 $dwordValue = 1
} else
{
	 $dwordValue = 0
}

set-itemproperty $key -name EnableLogging -value $dwordValue -type dword
set-itemproperty $key -name "65815" -value $dwordValue -type dword
set-itemproperty $key -name "65816" -value $dwordValue -type dword

