Param(	[string]$Seed     = 'Root'	,	`
		[int]$LengthAlpha = 8		,	`
		[int]$LengthOther = 3		,	`
		[int]$Count       = 10			`
		)
CLS
Get-Date

For(	$int1G = 1			;	`
		$int1G -LE $Count	;	`
		$int1G ++				`
		)
{	$strPwdG =	[system.web.security.membership]::	`
				GeneratePassword( $LengthAlpha, $LengthOther )
	Write-Host "  $strPwdG"
}#Next

"`r`nDone..."