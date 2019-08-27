### Numlock Script ###
## Checks if numlock is off, if its off it turns it on

function testnumlock
{
	if (-not [console]::NumberLock)
	{
		$w = New-Object -ComObject WScript.Shell;
		$w.SendKeys('{NUMLOCK}');
	}
}
testnumlock
