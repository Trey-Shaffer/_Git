##############################################################################
Function GetPSLibraryVersion-(){ Return '2.2.20160718' }
# 2.0 - First revision for module utilization.
# 2.1 - Corrected intra-module calls using GetVariable
# 2.2 - Added ErrorFormatter-
##############################################################################


<#############################################################################
.Synopsis
Returns Date and/or time in a compressed string.

.Description
Date format is YYYYMMDD.  Time format is HHMMSSmmm.
#>
Function DateTimeString-()
{#############################################################################
Param(	[dateTime]$DateTime = ( Get-Date )							,	`
		[ValidateSet( '_', '-', ' ' )] [string]$Separator  = '_'	,	`
		[switch]$Date												,	`
		[switch]$Time												,	`
		[switch]$Milliseconds											`
		)
	If( ( $Date -eq $False ) -AND ( $Time -eq $False ) ){ Return '' }

	$strDateStringL = ''
	$strTimeStringL = ''

	If( $Date )
	{	$strDateStringL = [string]($DateTime).Year

		$strTempL       = [string]($DateTime).Month
		If( $strTempL.Length -eq 1 ){ $strTempL = "0$strTempL" }
		$strDateStringL += $strTempL

		$strTempL = [string]($DateTime).Day
		If( $strTempL.Length -eq 1 ){ $strTempL = "0$strTempL" }
		$strDateStringL += $strTempL
	}#End If - Date

	If( $Time )
	{	$strTempL = [string]($DateTime).Hour
		If( $strTempL.Length -eq 1 )
		{	$strTempL = "0$strTempL" }
		$strTimeStringL = $strTempL

		$strTempL = [string]($DateTime).Minute
		If( $strTempL.Length -eq 1 )
		{	$strTempL = "0$strTempL" }
		$strTimeStringL += $strTempL

		$strTempL = [string]($DateTime).Second
		If( $strTempL.Length -eq 1 )
		{	$strTempL = "0$strTempL" }
		$strTimeStringL += $strTempL

		If( $Milliseconds )
		{	$strTempL = [string]($DateTime).Millisecond
			If( $strTempL.Length -eq 1 ){ $strTempL = "00$strTempL" }
			If( $strTempL.Length -eq 2 ){ $strTempL = "0$strTempL" }
			$strTimeStringL += $strTempL
		}#End If - Milliseconds
	}#End If - Time

	If( ( $Date -AND $Time ) )
	{ Return $strDateStringL + $Separator + $strTimeStringL }
	Else
	{ Return $strDateStringL + $strTimeStringL }
}#End Date Time String #######################################################


<#############################################################################
.Synopsis
Works with text encoded with the coresponding Encrypt-() function.

.Description
Reverses simple obfuscation to reconstruct clear text.
#>
Function Decrypt-()
{#############################################################################
Param(	[string]$CipherText	,	`
		[string]$Key			`
		)
	$arrTextL  = @()
	$arrInputL = $CipherText.ToCharArray()
	$strItemL  = ''
	$strOutL   = ''
	ForEach( $strCharL IN $arrInputL )
	{	$strItemL += $strCharL
		If( $strItemL.Length -eq 3 )
		{	$arrTextL += $strItemL
			$strItemL = ''
		}#End If
	}#Next

	$intKeyCounterL = 0
	ForEach( $intCharL IN $arrTextL )
	{	$strKeyCharL   = $Key.Substring( $intKeyCounterL, 1 )
		$intClearCharL = $intCharL - [Int][Char]$strKeyCharL
		$strOutL      += [Char]$intClearCharL
		$intKeyCounterL ++
		If( $intKeyCounterL -eq $Key.Length ){ $intKeyCounterL = 0 }
	}#Next
	Return $strOutL
}#End Decrypt ################################################################


<#############################################################################
.Synopsis
Simple obfuscation.

.Description
Creates three-digit representations of each character of input text.
Reverse with coresponding Decrypt-() function.
#>
Function Encrypt-()
{#############################################################################
Param(	[string]$ClearText	,	`
		[string]$Key			`
		)
	$arrTextL       = $ClearText.ToCharArray()
	$intKeyCounterL = 0
	ForEach( $strCharL IN $arrTextL )
	{	$strKeyCharL   = $Key.Substring( $intKeyCounterL, 1 )
		$intCryptCharL = [Int][Char]$strCharL + [Int][Char]$strKeyCharL
		$strCharOutL   = [string]$intCryptCharL

		While( $strCharOutL.Length -lt 3 ){ $strCharOutL = '0' + $strCharOutL }
		$strOutL += $strCharOutL
		$intKeyCounterL ++
		If( $intKeyCounterL -eq $Key.Length ){ $intKeyCounterL = 0 }
	}#Next
	Return $strOutL
}#End Encrypt ################################################################


<#############################################################################
.Synopsis
Adds/verifies last character of string.

.Description
Typically used to confirm backslash in path string, hence the default value
  of -EndChar argument is '\'.
No printed outputs, no strLogFileFqG.
#>
Function EndChar-()
{#############################################################################
Param(	[string]$String			,	`
		[string]$EndChar = '\'		`
		)
	If( $String.EndsWith( $EndChar ) ){ Return $String }
	Else{ Return $String + $EndChar }
}#End End Char ###############################################################


<#############################################################################
.Synopsis
Generates a verbose string representing duration between two dateTime values.

.Description
-Stop defaults to NOW, so single argument is assumed to be -Start.
#>
Function EtString-()
{#############################################################################
Param(	[Parameter( Mandatory = $True ) ][DateTime]$StartTime	,	`
		[DateTime]$StopTime = (Get-Date)							`
		)
	$strElapsedTimeL = ""
	$strTempL        = ""
	$ElapasedL       = New-TimeSpan $StartTime $StopTime

	#Build time string...
	If( $ElapasedL.Days -gt 0 )
	{	$strTempL = [string]$ElapasedL.Days
		$strElapsedTimeL += "$strTempL day(s)"
	}#End If - Days

	If( $ElapasedL.Hours -gt 0 )
	{	If( $strElapsedTimeL.Length -gt 0 )
		{	$strElapsedTimeL += ", "
		}#End If
		$strTempL = [string]$ElapasedL.Hours
		$strElapsedTimeL += "$strTempL hours(s)"
	}#End If - Hours

	If( $ElapasedL.Minutes -gt 0 )
	{	If( $strElapsedTimeL.Length -gt 0 )
		{	$strElapsedTimeL += ", " }
		$strTempL = [string]$ElapasedL.Minutes
		$strElapsedTimeL += "$strTempL minute(s)"
	}#End If - Minutes

	If( $ElapasedL.Seconds + $ElapasedL.Milliseconds -gt 0 )
	{	$strMSecsL = [string]$ElapasedL.Milliseconds
		$strMSecsL = Padstring- $strMSecsL 3 -PadChar "0" -Left

		If( $strElapsedTimeL.Length -gt 0 )
		{	$strElapsedTimeL += ", " }
		$strTempL = [string]$ElapasedL.Seconds	+	`
					"."							+	`
					$strMSecsL
		$strElapsedTimeL += "$strTempL seconds"
	}#End If - Seconds
	Return $strElapsedTimeL
}#End ET String ##############################################################


<############################################################################
.Synopsis
Formats a standard error object for display.

.Description
Accepts an error object, typically $Error[0], after a runtime error occurs
  and is recognized.  Optional description string is a free form description
  probably indicating what script activity failed.
#>
Function ErrorFormatter-()
{############################################################################
[CmdletBinding()]
Param(	$Error					,	`
		$Description   = $Null	,	`
		$strLogFileFqG = $Null	,	`
		[switch]$Quiet				`
		)
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$strExMessageL    = $Error.Exception.Message
	$strExMessageL    = "Exception: $strExMessageL"
	$strExMessageL    = Paragraph- $strExMessageL

	$strCategoryInfoL = $Error.CategoryInfo
	$strCategoryInfoL = " Category: $strCategoryInfoL"
	$strCategoryInfoL = Paragraph- $strCategoryInfoL

	$strFqErrorIdL    = $Error.FullyQualifiedErrorID
	$strFqErrorIdL    = "  Full ID: $strFqErrorIdL"
	$strFqErrorIdL    = Paragraph- $strFqErrorIdL

	$strScriptNameL   = $Error.InvocationInfo.ScriptName
	$strScriptNameL   = "   Script: $strScriptNameL"
	$strScriptNameL   = Paragraph- $strScriptNameL

	$strLineInfoL     =	'Line/Char: ' 									+	`
						[string]$Error.InvocationInfo.ScriptLineNumber	+	`
						' / '											+	`
						$Error.InvocationInfo.OffsetInLine
	$strLineInfoL     = Paragraph- $strLineInfoL

	$strInvocInfoL    = ( $Error.InvocationInfo.Line ).Trim()
	$strInvocInfoL    = "  Command: $strInvocInfoL"
	$strInvocInfoL    = Paragraph- $strInvocInfoL

	$strMsgL =	"  **An error has occurred.`r`n"	+	`
				$strExMessageL						+	`
				$strCategoryInfoL					+	`
				$strFqErrorIdL						+	`
				$strScriptNameL						+	`
				$strLineInfoL						+	`
				$strInvocInfoL

	If( $Description )
	{	$strDescriptionL  = "     Note: $Description"
		$strDescriptionL  = Paragraph- $strDescriptionL
		$strMsgL         += $strDescriptionL
	}#End If
	
	If( $Quiet -eq $False )
	{	WriteBoth- $strMsgL -Red
		$strMsgL | Out-File -FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
	}#End If - quiet
	Return $strMsgL
}#End Error Formatter ########################################################


##############################################################################
Function ExitProcessing-()
{#############################################################################
[CmdletBinding()]
Param(	[int]$ReturnCode    = 0		,	`
		[dateTime]$StartTime		,	`
		[dateTime]$StopTime			,	`
		$strLogFileFqG      = $Null		`
		)
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - null

	If( -NOT $StopTime ){ $StopTime = Get-Date }

	$strFinishTimeL = $StopTime.ToString()
	#Get calling script start time and log file info.
	If( -NOT $StartTime )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'dteStartG'
	}#End If
	If( -NOT $strLogFileFqG ){ $strLogFileFqG = 'No log specified' }

	#Confirm Start time.
	If( -NOT $StartTime )
	{	If( -NOT $dteStartG )
		{	$strMsgL =	'**No start time available for ET calculation.'
			Write-Verbose -Message $strMsgL
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
			$StartTime = Get-Date
		}#End If
		Else { $StartTime = [dateTime]$dteStartG }
	}#End If - start null
	Else { $StartTime = [dateTime]$StartTime }

	$strEtL = EtString- -Start $StartTime  -Stop $StopTime

	$strMsgL =	"`r`n"						+	`
				('=' * 78) + "`r`n"			+	`
				"Processing complete...`t"	+	`
				$strFinishTimeL				+	`
				"`r`n  Duration: "			+	`
				$strEtL						+	`
				"`r`n  Returning value: "	+	`
				$ReturnCode					+	`
				".`r`n`r`n"

	If( $strLogFileFqG -ne '' )
	{	$strMsgL +=	'  Detailed information may be found in log file:'	+	`
					"`r`n`t"											+	`
					$strLogFileFqG										+	`
					"`r`n`r`n`r`n"
	}#End If
	Write-Output $strMsgL
	$strMsgL | Out-File -FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append
	Exit $ReturnCode
}#End Exit Processing ########################################################


<#############################################################################
.Synopsis
Converts a numeric value, typically a file size to a formatted string.

.Description
The combination of -Size and -Unit are used to calculate a size in bytes.
Then that value is formatted with commas and decimals and returned in a
  string padded to the non-zero value of the -Length parameter, if the
  -Length parameter is greater than the length of the formatted string.
#>
Function FileSize-()
{#############################################################################
Param(	[int64]$Size = 1024							,	`
		[ValidateSet(	'B', 'KB', 'MB', 'TB' ) ]		`
		[string]$Unit   = 'B'						,	`
		[int]$Precision = 2							,	`
		[int]$Length    = 0								`
		)
	Switch ( $Unit )
	{	B  { $Size = $Size; Break }
		KB { $Size = $Size * 1KB; Break }
		MB { $Size = $Size * 1MB; Break }
		GB { $Size = $Size * 1GB; Break }
		TB { $Size = $Size * 1TB; Break }
		Default { Write-Output '**Error' }
	}#End switch

	If( $Size -ge 1TB )
	{	$strReturnL = "{0:N$Precision}" -f ( $Size / 1TB )
		$strReturnL += ' TB'
	}#End If - TB
	ElseIf( $Size -ge 1GB )
	{	$strReturnL = "{0:N$Precision}" -f ( $Size / 1GB )
		$strReturnL += ' GB'
	}#End GB
	ElseIf( $Size -ge 1MB )
	{	$strReturnL = "{0:N$Precision}" -f ( $Size / 1MB )
		$strReturnL += ' MB'
	}#End MB
	ElseIf( $Size -ge 1024 )
	{	$strReturnL  = "{0:N$Precision}" -f ( $Size / 1KB )
		$strReturnL += ' KB'
	}#End KB
	Else
	{	$strReturnL  = "{0:N0}" -f $Size
		$strReturnL += ' Bytes'
	}#End default

	If( $Length -gt $strReturnL.Length )
	{	$strReturnL = PadString- $strReturnL -Length $Length }
	Return $strReturnL
}#End - File Size ############################################################


<#############################################################################
.Synopsis
Gets variables defined in the calling routine.

.Description
Originally from:
  https://gallery.technet.microsoft.com/Inherit-Preference-82343b9d/file/107568/9/Get-CallerPreference.ps1

This version reduced to only get specified variables rather than all the
  defined system state variables.
There is a gotcha...  This works one way in PowerGUI and PowerShell ISE,
  and another way in a native PowerShell session.
In a native session this only returns variables from the scope of the calling
  routine, hence it is limited to one level.
If running in the ISE, it somehow supported nested layers; in other words,
  WriteLog- worked when it was called from WriteBoth-, but WriteLog- could not
  get the strLogFileFqG value when running in a native session.
#>
Function GetCallerVariable-()
{#############################################################################
[CmdletBinding(DefaultParameterSetName = 'AllVariables')]
Param(	[ Parameter( Mandatory = $True ) ]										`
		[ ValidateScript( {	$_.GetType().FullName -eq							`
							'System.Management.Automation.PSScriptCmdlet'		`
							}													`
						)														`
		]																		`
		$Cmdlet																,	`
		[ Parameter( Mandatory = $True ) ]										`
		[System.Management.Automation.SessionState]								`
		$SessionState														,	`
		[ Parameter(	ParameterSetName  = 'Filtered'						,	`
						ValueFromPipeline = $True								`
						)														`
		]																		`
		[ Parameter( Mandatory = $True ) ]										`
		[string[]]$VarNames														`
		)
	Begin{ $htFilterHashL = @{} }
	Process
	{	If( $VarNames )
		{	ForEach( $VarNameL IN $VarNames )
			{	$htFilterHashL[ $VarNameL ] = $True }
		}#End If
	}#End Process

	End
	{	ForEach ( $VarNameL in $VarNames )
		{	$oVariableL = $Cmdlet.SessionState.PSVariable.Get($VarNameL)
			If( $oVariableL )
			{	If ($SessionState -eq $ExecutionContext.SessionState)
				{	Set-Variable	-Scope 1					`
									-Name  $oVariableL.Name		`
									-Value $oVariableL.Value	`
									-Force						`
									-Confirm:$False				`
									-WhatIf:$False
				}#End If - $SessionState
				Else
				{ $SessionState.PSVariable.Set( $oVariableL.Name, $oVariableL.Value ) }
			}#End If - $oVariableL
		}#Next
	}#End - End
}#End  Get Caller Variable 2 #################################################


<#############################################################################
.Synopsis
Returns size of folder in bytes.

.Description
Accepts UNC or local path.  Recurse option includes subfolders.
Test option verifies folder exists before attempting retrieval.
#>
Function GetFolderSize-()
{#############################################################################
[CmdletBinding()]
Param(	[string]$UNC			,	`
		[switch]$Recurse		,	`
		[switch]$Test			,	`
		$strLogFileFqG = $Null		`
		)
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null

	If( $Test )
	{	If( -NOT( Test-Path -Path $UNC ) )
		{	$strMsgL =	'**Error - Invalid UNC Path'	+	`
						"`r`n  "						+	`
						$UNC							+	`
						', not found.'
			Write-Output $strMsgL	-Foreground 'Red'
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
			ExitProcessing- -strLogFileFqG $strLogFileFqG
		}#End If - Test path
	}#End If - Test

	If( $Recurse ){ $colFilesL = Get-ChildItem $UNC -Recurse }
	Else{  $colFilesL = Get-ChildItem $UNC }

	$dblTotalSizeL = 0
	ForEach( $oFileL IN $colFilesL )
	{	$strAttributesL = $oFileL.Attributes.ToString().ToLower()

		If( $strAttributesL.Contains( 'directory' ) )
		{	$strMsgL = '  Folder - ' + $oFileL.Name
			Write-Verbose -Message $strMsgL
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
		}#End If
		Else
		{	$dblTotalSizeL += $oFileL.Length
			$strMsgL =	'  File - '			+	`
						$oFileL.FullName	+	`
						' - '				+	`
						$oFileL.Length
			Write-Verbose -Message $strMsgL
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
		}#End If
	}#Next - File
	Return $dblTotalSizeL
}#End - Get Folder Size ######################################################


<#############################################################################
.Synopsis
Takes a string argument and examines it to see if it is single item,
  a comma-separated list, or a file name.  Then creates an array from the
  item, comma-separated list or contents of the file.

.Description
This calls the ReadInputTextFile-() function, but does not offer the option
  of returning a hashtable, which can be done calling the function directly.
Local path of calling script required.
#>
Function GetList-()
{#############################################################################
[CmdletBinding()]
[OutputType("System.Object[]")]
Param(	[string]$List					,	`
		[string]$Delimiter   = ','		,	`
		[string]$DefaultPath 			,	`
		$strLogFileFqG       = $Null		`
		)
	$dteStartL = Get-Date
	
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$arrListL = @()
	$DefaultPath = EndChar- $DefaultPath -EndChar '\'

	If( $List.Contains( '.' ) )
	{	$arrTempL = $List.Split( '.' )
		If( $arrTempL.GetUpperBound(0) -eq 0 )
		{	#Only one period.
			#Assume it is a file name.
			If( -NOT( $List.Contains( '\' ) ) )
			{	#File name only, not full path.
				$List = $DefaultPath + $List
			}#End If - path
		}#End If - file name
		$arrListL = ReadInputTextFile-	-InputFile     $List			`
										-strLogFileFqG $strLogFileFqG
	}#End If - .
	Else
	{	#List is a delimited list of items
		$arrListL = $List.Split( $Delimiter )
	}#End Else

	$strMsgL = '  The following items were retrieved:'
	ForEach( $strItemL IN $arrListL ){ $strMsgL += "`r`n    $strItemL" }
	Write-Verbose -Message $strMsgL
	$strMsgL | Out-File	-FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append
	
	$strMsgL = EtString- -StartTime $dteStartL
	$strMsgL = "  List generation, $strMsgL."
	Write-Verbose $strMsgL
	
	Return ,$arrListL
}#End Get List ###############################################################


<#############################################################################
.Synopsis
Wrapper for GetList-.  Adds the option to read the value of 'Database', in
  which case the list of server names is read from the Script-scope database.

.Description
Just checks for database option and if not found calls Get-List-.
#>
Function GetServerList-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Object[]" )]
Param(	[string]$List			,	`
		[switch]$NetBIOS		,	`
		$Connection    = $Null	,	`
		$strLogFileFqG = $Null		`
		)
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - strLogFileFqG

	$arrListL = @()
	If( $List )
	{	If( $List -eq 'Database' )
		{	#Get list from database
			#Not implemented...
			If( -NOT $Connection )
			{	GetCallerVariable-												`
								-Cmdlet       $PSCmdlet							`
								-SessionState $ExecutionContext.SessionState	`
								-VarNames     'oDbConnectionG'
				$Connection = $oDbConnectionG
			}#End If - connection

			$strQueryL    = "Select * From sys_Servers"
			$oServerDataL = SqlInitializeDataset-							`
											-Query      $strQueryL			`
											-Connection $Connection			`
											-strLogFileFqG $strLogFileFqG

			ForEach( $oServerL in $oServerDataL.Tables[0].Rows )
			{	$strServerNameL  = [string]$oServerL.ServerName
				$arrListL += $strServerNameL
			}#Next - oServerL
		}#End If - Database
		ElseIf( $List.Contains( '=' ) )
		{	#Assume distinguishedName for LDAP lookup.
			WriteVerbose- -InputText '  LDAP root path given.'
			$colComputersL = Get-ADComputer -SearchBase $List -Filter *
			$strMsgL       = '  Computers from Active Directory:'
			ForEach( $oComputerL IN $colComputersL )
			{	$arrListL += $oComputerL.Name }
		}#End ElseIf - distinguishedName
		Else
		{	#Not calling database, process as list/file
			$arrListL = GetList-	-List          $List			`
									-strLogFileFqG $strLogFileFqG
		}#End Else - not database
	}#End If - List not null
	Else
	{	$strMsgL = 'No server name/list specified.  Using local computer.'
		Write-Output $strMsgL
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding   'ASCII'			`
							-Append
		$arrListL = @( $Env:ComputerName )
	}#End Else
	
	#If requested, remove potential domain names.
	If( $NetBIOS )
	{	$colNetBiosL = @()
		ForEach( $strItemL IN $arrListL )
		{	$arrTempL = $strItemL.Split( '.' )
			$colNetBiosL += ,$arrTempL[0]
		}#Next - strItemL
		$arrListL = $colNetBiosL
	}#End If - NetBIOS
	
	$strMsgL = '  The following server(s) will be processed:'
	ForEach( $strServerL IN $arrListL )
	{	$strMsgL += "`r`n    $strServerL" }
		Write-Verbose -Message $strMsgL
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
	$arrListL = $arrListL | Sort-Object
	
	Return ,$arrListL
}#End Get Server List ########################################################


<#############################################################################
.Synopsis
Creates an object representing SMTP server, for use sending email.

.Description
Creates a script-scope object, oSmtpServerG with attributes:
  ServerName
  Port
  UseAuth
  AuthName
  AuthPwd

Configurations are stored in table sys_SystemInformation, and referenced by
  their key value, ConfigID
#>
Function GetSmtpServer-()
{#############################################################################
[CmdletBinding()]
Param(	$Connection         = $Null	,	`
		[int]$ConfigNumber  = 1		,	`
		$strLogFileFqG      = $Null		`
		)
	If( -NOT $Connection )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'oDbConnectionG'
		$Connection = $oDbConnectionG
	}#End If - connection null
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$strQueryL   =	'Select * From sys_SystemInformation '	+	`
					"WHERE ConfigID = $ConfigNumber"
	$oSqlReaderL =	SqlInitializeDataReader-	-Query         $strQueryL		`
												-Connection    $Connection		`
												-strLogFileFqG $strLogFileFqG
	While( $oSqlReaderL.Read() )
	{	#Create SMTP Server Object
		$oSmtpServerL = New-Object -TypeName System.Object
		Add-Member	-InputObject $oSmtpServerL							`
					-Type        NoteProperty							`
					-Name        ServerName								`
					-Value       $oSqlReaderL['SmtpServer'].ToString()
		Add-Member	-InputObject $oSmtpServerL							`
					-Type        NoteProperty							`
					-Name        Port									`
					-Value       $oSqlReaderL['SmtpPort']
		Add-Member	-InputObject $oSmtpServerL							`
					-Type        NoteProperty							`
					-Name        UseAuth								`
					-Value       $oSqlReaderL['SmtpUseAuth']
		Add-Member	-InputObject $oSmtpServerL							`
					-Type        NoteProperty							`
					-Name        AuthName								`
					-Value       $oSqlReaderL['SmtpAuthName'].ToString()
		Add-Member	-InputObject $oSmtpServerL							`
					-Type        NoteProperty							`
					-Name        AuthPwd								`
					-Value       $oSqlReaderL['SmtpAuthPwd'].ToString()
		Add-Member	-InputObject $oSmtpServerL							`
					-Type        NoteProperty							`
					-Name        EmailTo								`
					-Value       $oSqlReaderL['EmailTo'].ToString()
		Add-Member	-InputObject $oSmtpServerL							`
					-Type        NoteProperty							`
					-Name        EmailFrom								`
					-Value       $oSqlReaderL['EmailFrom'].ToString()
		Add-Member	-InputObject $oSmtpServerL							`
					-Type        NoteProperty							`
					-Name        Subject								`
					-Value       $oSqlReaderL['Subject'].ToString()
	}#WEnd
	$oSqlReaderL.Close()
	$strMsgL = '  SMTP Server object initialized.'
	Write-Verbose -Message $strMsgL
	Return $oSmtpServerL
}#End Get SMTP Server ########################################################


<#############################################################################
.Synopsis
Initializes logging based on script's command line parms.

.Description
  -LogAppend controls overwrite of existing file of same name.
  -LogDate, -LogTime control date/timestamp in file name.
This function does NOT actually create/initialize the log file. (bad name)
This function builds/returns the log file name and deletes an existing file,
  of the same name, if it exists, and the APPEND option is not specified.
Call to this function is usually followed by call to WriteBanner-
#>
Function InitializeLog-()
{#############################################################################
[CmdletBinding()]
Param(	[string]$LogBase    = 'Log'		,	`
		[string]$LogFolder  = ''		,	`
		[string]$LogExt     = '.txt'	,	`
		[bool]$LogAppend    = $False	,	`
		[bool]$LogDate      = $False	,	`
		[bool]$LogTime      = $False	,	`
		$strScriptPathG     = $Null		,	`
		$strLogFileFqG      = $Null
		)
	If( -NOT $strScriptPathG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG,strScriptPathG'
	}#End If - null

	$strDateStringL = ''
	If( $LogFolder -eq '' ){ $LogFolder = $strScriptPathG }
	$LogFolder = EndChar- $LogFolder -EndChar '\'

	If( -NOT( Test-Path -Path $LogFolder ) )
	{	Write-Output	'**Error - Log folder is not valid.'	`
					-Foreground 'Red'
		ExitProcessing- -ReturnCode 77					`
						-strLogFileFqG $strLogFileFqG
	}#End If

	#Construct date/time portion of file name
	If($LogDate -AND $LogTime )
	{	$strDateStringL = DateTimeString- -Date -Time }
	Else
	{	If( $LogDate ){ $strDateStringL = DateTimeString- -Date }
		If( $LogTime ){ $strDateStringL = DateTimeString- -Time }
	}#End Else - date + time

	$strLogFileFqL =	$LogFolder		+	`
						$LogBase		+	`
						$strDateStringL	+	`
						$LogExt
	#Delete existing log file, if needed
	If( -NOT $LogAppend )
	{	If( Test-Path -Path( $strLogFileFqL ) )
		{	Remove-Item $strLogFileFqL }
	}#End If - Test Path

	Return $strLogFileFqL
}#End Initialize Log #########################################################


##############################################################################
Function IsAdmin-()
{#############################################################################
	$UserL = [Security.Principal.WindowsIdentity]::GetCurrent()
	(New-Object -TypeName Security.Principal.WindowsPrincipal $UserL).IsInRole(	`
			[Security.Principal.WindowsBuiltinRole]::Administrator)
}#End Is Admin ###############################################################


##############################################################################
Function IsDbNull-( $InputP )
{#############################################################################
	Return [System.DBNull]::Value.Equals($InputP)
}#End Is DbNull ##############################################################


<#############################################################################
.Synopsis
Formatted version of Get-Type.

.Description
Intended for debugging.
#>
Function ListObjectProperties-( $Object )
{#############################################################################
	$strMsgL = "`r`n  Type name is " + $Object.GetType().Fullname
	Write-Output $strMsgL

	$colPropertiesL = $Object.PsObject.Properties

	If( !( $colPropertiesL ) )
	{	Write-Output "  Object has no properties"
		Return
	}#End If

	$intPropertyIndexL = 0

	#Prepare formatting
	$strFormatL = "  {0}) Name: {1} ( {2} )"

	Foreach ( $oPropertyL in $colPropertiesL)
	{	$intPropertyIndexL ++
		$oPropertyNameL = $oPropertyL.Name

		#place variable name in single quotes to ensure that
		#  PowerShell does not evaluate\substitute value
		$oPropertyNameFullL = '$Object' + '.' + $oPropertyNameL

		#prepare to use variable substitution
		# Invoke-Expression
		# http://technet.microsoft.com/en-us/library/dd347550.aspx
		$oPropertyValueL = invoke-expression $oPropertyNameFullL

		$strMsgL = [String]::Format(	$strFormatL			,	`
										$intPropertyIndexL	,	`
										$oPropertyNameL		,	`
										$oPropertyValueL		`
										)
		Write-Output $strMsgL
	}#Next
}#End List Object Properties #################################################


##############################################################################
Function LoadEx2007Snapin-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Boolean" )]
Param(	[switch]$NoAbort )
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$strSnapinL = 'Microsoft.Exchange.Management.PowerShell.Admin'
	Add-PsSnapin $strSnapinL -EA 'SilentlyContinue'
	$oSnapinL = Get-PSSnapin $strSnapinL
	If( $oSnapinL ){ Return $True }
	Else
	{	$strMsgL =	'**Error - Unable to load Exchange 2007 snapin.'
		Write-Output $strMsgL -Foreground 'Red'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding   'ASCII'			`
							-Append
		If( $NoAbort ){ Return $False }
		Else { ExitProcessing- -strLogFileFqG $strLogFileFqG }
	}#End If
}# End Load Exchange 2007 Snapin #############################################


##############################################################################
<#
.Synopsis
Checks to see if module loaded.  If not, attempt to load.  If fail, attempt
  to import and load it.
.Description
Checks to see if specified module is loaded.
If not, checks to see if it is registered as "Available".
If not, checks to see if it exists in the local profile, \Modules path.
If not, attempts to copy it from the specified source location and then
  import it.
.Example
LoadModule- -ModuleName 'poaShell'  -ModulePath 'C:\Temp'
	Checks to see if poaShell is loaded/available and attempts to copy
	  and load it if not.  Aborts on failure.
#>
Function LoadModule-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Boolean" )]
Param(	#Name of module to loaded									`
		[Parameter ( Mandatory = $True ) ]							`
		[string]$ModuleName										,	`
		#Path to source holding module to install and load			`
		[string]$ModulePath = $Null								,	`
		#Switch to only check load status no attempt to load		`
		[switch]$CheckOnly										,	`
		#Switch to allow continue after failure to load				`
		[switch]$NoAbort										,	`
		$strLogFileFqG = $Null										`
		)
	If( -NOT $ModulePath )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'ModulePath'
	}#End If - ModulePath
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$strMsgL = "  Examining status of module: $ModuleName..."
	Write-Verbose -Message $strMsgL
	$strMsgL | Out-File	-FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append

	$blnIsAvailableL   = $False
	$blnIsLoadedL      = $False

	#Get modules already loaded.
	$colLoadedModulesL = Get-Module

	If( $colLoadedModulesL )
	{	#Some are loaded.  Check to see if target is loaded.
		If( $colLoadedModulesL -IsNot [array] )
		{	$colLoadedModulesL = @( $colLoadedModulesL ) }

		ForEach( $oLoadedModuleL IN $colLoadedModulesL )
		{	If( $oLoadedModuleL.Name -eq $ModuleName )
			{	$blnIsLoadedL = $True
				Write-Verbose -Message $strMsgL
				$strMsgL | Out-File	-FilePath $strLogFileFqG	`
									-Encoding 'ASCII'			`
									-Append
				Return $True
			}#End If - found
		}#Next - oLoadedModuleL

		#Did not find it in loaded modules.
		If( $CheckOnly )
		{	#Not attempting to load.
			Return $False
		}#End If - CheckOnly
	}#End If - some module loaded

	If( $blnIsLoadedL -eq $False )
	{	#Not loaded, check to see if registered but not loaded
		$strMsgL = '  Module is not loaded.  Checking available...'
		Write-Verbose -Message $strMsgL
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		$colLoadedModulesL = Get-Module -ListAvailable
		If( $colLoadedModulesL )
		{	ForEach( $oLoadedModuleL IN $colLoadedModulesL )
			{	$strMsgL =	'  Loaded module: '	+	`
							$oLoadedModuleL.Name
				Write-Verbose -Message $strMsgL
				$strMsgL | Out-File	-FilePath $strLogFileFqG	`
									-Encoding 'ASCII'			`
									-Append
				If( $oLoadedModuleL.Name -eq $ModuleName )
				{	#Flag as available, load later...
					$blnIsAvailableL = $True
					Break
				}#End If - found
			}#Next - oLoadedModuleL

			If( $blnIsAvailableL -eq $False )
			{	#Module was not loaded AND is not available.
				If( $ModulePath = $Null )
				{	#No source path was specified to install module.
					$strMsgL =	'**Warning - specified module not '	+	`
								'loaded OR available.'
					Write-Output $strMsgL -Foreground 'Yellow' -Background 'Black'
					$strMsgL | Out-File	-FilePath $strLogFileFqG	`
										-Encoding   'ASCII'			`
										-Append
					Return $False
				}#End If - no ModulePath
			}#End If - not available
		}#End If - some module available
	}#End If - Available

	If( -NOT( $blnIsAvailableL ) )
	{	$strMsgL = "`r`n  Module is not available"
		Write-Verbose -Message $strMsgL
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append

		#Now look at profile path, then \Modules to see if exists in
		#  local profile.
		#
		#Get profile path.
		$strProfileL = $Profile		#Includes profile file name
		$arrTempL    = $strProfileL.Split( '\' )

		#Build path to profile folder
		$strProfileFolderL = ''
		$arrPathL          = $strProfileL.Split( '\' )
		For(	$int1L = 0								;	`
				$int1L -lt $arrPathL.GetUpperBound(0)	;	`
				$int1L ++									`
			)
		{	$strProfileFolderL += $arrPathL[ $int1L ] + '\' }

		$strMsgL =	"`r`n  Local profile folder exists:`r`n`t$strProfileFolderL"
		Write-Verbose -Message $strMsgL
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		#Modules folder may not exist
		#Not created until first module installed/loaded
		$strModulesPathL   = $strProfileFolderL + 'Modules\'

		#Check for Modules Folder
		If( Test-Path -Path $strModulesPathL )
		{	$strMsgL = "`r`n  Local Modules folder exists"
			Write-Verbose -Message $strMsgL
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
		}#End If - Test-Path
		Else
		{	$strMsgL = "`r`n  Modules folder does not exist.  Creating it..."
			Write-Verbose -Message $strMsgL
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
			}#End If - no modules folder

		#Target folder for specified module.
		$strTargetFolderL = $strModulesPathL + $ModuleName

		#Check for target module folder
		If( Test-Path -Path $strTargetFolderL )
		{	$strMsgL =	"`r`n  Local Module folder exists: "	+	`
						"`r`n`t$strTargetFolderL"
			Write-Verbose -Message $strMsgL
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
			#Folder exists.  Assume module in folder.
			$blnIsAvailableL = $True
		}#End If - Test-Path
		Else
		{	#Folder for module is not found in profile.
			#See if a module source location is specified.
			If( $ModulePath )
			{	#A source folder for the module is specified.
				#Module will be copied from there to local profile.
				$strSourceFolderL  = EndChar- $ModulePath
				$strSourceFolderL += $ModuleName
				$strMsgL =	"`r`n  Local module folder "	+	`
							"does not exist: `r`n`t"		+	`
							$strTargetFolderL				+	`
							"`r`n  Will copy from source,"	+	`
							"`r`n`t$strSourceFolderL"
				Write-Verbose -Message $strMsgL
				$strMsgL | Out-File	-FilePath $strLogFileFqG	`
									-Encoding 'ASCII'			`
									-Append
				If( Test-Path -Path $strSourceFolderL )
				{	#Source folder exists.  Assume module exists in folder.
					Copy-Item	-Path        $strSourceFolderL	`
								-Destination $strModulesPathL	`
								-EA          'SilentlyContinue'	`
								-Recurse						`
								-Force
					If( $? )
					{	$blnIsAvailableL = $True }
					Else
					{	$strMsgL =	'**Error - Source module not found'	+	`
									"`r`n`t$strSourceFolderL"
						Write-Output $strMsgL -Foreground 'Red'
						$strMsgL | Out-File	-FilePath $strLogFileFqG	`
											-Encoding 'ASCII'			`
											-Append

						$strErrL = "Failed to copy module, $ModuleName"
						ErrorFormatter-	-Error         $Error[0]		`
										-Description   $strErrL			`
										-strLogFileFqG $strLogFileFqG

						If( $NoAbort ){ Return $False }
						Else
						{	ExitProcessing-	-ReturnCode    78				`
											-strLogFileFqG $strLogFileFqG	`
						}#End Else
					}#End If - error
				}#End If - source path exists
				Else
				{	$strMsgL =	'**Error - Source path for module '	+	`
								'installation does not exist'		+	`
								"`r`n`t$strSourceFolderL"
					Write-Output $strMsgL -Foreground 'Red'
					$strMsgL | Out-File	-FilePath $strLogFileFqG	`
										-Encoding 'ASCII'			`
										-Append
					ExitProcessing-	-ReturnCode    79				`
									-strLogFileFqG $strLogFileFqG
				}#End Else - source not found
			}#End Else - Module now copied to destination
			Else
			{	$strMsgL =	'**Error - Specified PS module not '	+	`
							'installed or available locally'		+	`
							"`r`n`tModule: $ModuleName"				+	`
							"`r`n  No alternate source for "		+	`
							'installation was specified.'
				Write-Output $strMsgL -Foreground 'Red'
				$strMsgL | Out-File	-FilePath $strLogFileFqG	`
									-Encoding 'ASCII'			`
									-Append
				ExitProcessing-	-ReturnCode    79				`
								-strLogFileFqG $strLogFileFqG
			}#End Else - Module source not specified.
		}#End If - module source path specified
	}#End If - not avaialable
	Else
	{	#Module available, try to load it
		$strEaHoldL            = $ErrorActionPreference
		$ErrorActionPreference = 'SilentlyContinue'
		$strMsgL               = '  Loading module...'
		Write-Verbose -Message $strMsgL
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		Import-Module $ModuleName

		If( $? )
		{	Return $True }
		Else
		{	$strErrL = "Failed to load module, $ModuleName"
			ErrorFormatter-	-Error         $Error[0]		`
							-Description   $strErrL			`
							-strLogFileFqG $strLogFileFqG

			If( $NoAbort ){ Return $False }
			Else
			{	ExitProcessing-	-ReturnCode    77				`
								-strLogFileFqG $strLogFileFqG
			}#End Else
		}#End If - error
		$ErrorActionPreference = $strEaHoldL
	}#End Else - available
	Return $False	#Shouldn't actually get to here...
}# End Load Module ###########################################################


<#############################################################################
.Synopsis
Convert a MS-DOS format local path to a UNC format URI.

.Description
Required path in MS-DOS format, including drive letter, is converted to a UNC
  using ComputerName and the specified drive's administrative, $, share.
#>
Function MakeUNC-()
{#############################################################################
[CmdletBinding()]
Param(	$Path							,	`
		$Computer = $Env:ComputerName	,	`
		$blnLogInitializedG = $Null		,	`
		$strLogFileFqG      = $Null			`
		)
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$strPathL = $Path.ToString()
	$arrTempL = $strPathL.Split( '\' )
	If( -NOT( $arrTempL[0].Contains( ':' ) ) )
	{	$strMsgL =	'**Error - Invalid MS-DOS path'				+	`
					"`r`n"										+	`
					'  The input path, '						+	`
					$strPathL									+	`
					' should be in MS-DOS format.'				+	`
					"`r`n"										+	`
					'  It should begin with a drive letter '	+	`
					'and include a colon (:)'
		Write-Output $strMsgL -Foreground 'Red'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		ExitProcessing- -strLogFileFqG $strLogFileFqG
	}#End If - Error

	$strDriveL = $arrTempL[0].Replace( ':', '' )
	$strUncL   =	'\\'		+	`
					$Computer	+	`
					'\'			+	`
					$strDriveL	+	`
					'$'
	For(	$int1L = 1								;	`
			$int1L -le $arrTempL.GetUpperBound(0)	;	`
			$int1L ++									`
		)
	{	$strMsgL =	'  Item '			+	`
					$int1L				+	`
					': '				+	`
					$arrTempL[ $int1L ]
		Write-Verbose -Message $strMsgL
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		$strUncL += '\' + $arrTempL[ $int1L ]
	}#Next
	Return $strUncL
}#End - Make UNC #############################################################


<#############################################################################
.Synopsis
Creates a connection object to an Access database.

.Description
For backward compatability, tries a 32-bit driver connection first, then
  follows with a 64-bit attempt.
#>
Function MdbConnect-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Object[]" )]
Param(	[Parameter( Mandatory = $True ) ]		`
		$DB									,	`
		[Parameter( Mandatory = $True ) ]		`
		$Path								,	`
		$User          = "Admin"			,	`
		$Pwd           = ""					,	`
		$strLogFileFqG = $Null					`
		)
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$strTargetDbL = ( EndChar- -String $Path -EndChar -String '\' ) + $DB
	If( _NOT Test-Path -Path $strTargetDbL )
	{	$strMsgL = '**Error - specified database file not found.'
		Write-Output $strMsgL -Foreground 'Red'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		ExitProcessing- -strLogFileFqG $strLogFileFqG
	}#End If

	$strConnect32L	= "Driver={Microsoft Access Driver (*.mdb)};"
	$strConnect64L	= "Driver={Microsoft Access Driver (*.mdb, *.accdb)};"

	$Path = EndChar- -String $Path
	$strConnect2L =	"Dbq="			+	`
					$DB				+	`
					";DefaultDir="	+	`
					$Path			+	`
					";Uid="			+	`
					$User			+	`
					";Pwd="			+	`
					$Pwd			+	`
					";"
	Write-Verbose -Message $strConnect2L
	$strMsgL | Out-File	-FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append

	#Store Error action preference
	$strEarHoldL           = $ErrorActionPreference
	$ErrorActionPreference = "SilentlyContinue"

	oConnectionL = New-Object -ComObject ADODB.Connection

	#Attempt 32 bit connect first
	$strCStringL = $strConnect32L + $strConnect2L
	oConnectionL.Open( $strCStringL )
	If( -NOT $? )
	{	#Connection attempt failed, try 64 bit
		$strCStringL = $strConnect64L + $strConnect2L
		oConnectionL.Open( $strCStringL )
		If( -NOT $? )
		{	#Second attempt failed
			$strErrL =	"  Both 64 bit and 32 bit connection attempts "	+	`
						"to connect to database, "						+	`
						$DB												+	`
						", failed.  This is the error for the 64 bit "	+	`
						"attempt."
			ErrorFormatter-	-Error         $Error[0]		`
							-Description   $strErrL			`
							-strLogFileFqG $strLogFileFqG

			$strErrL =	"  Both 64 bit and 32 bit connection attempts "	+	`
						"to connect to database, "						+	`
						$DB												+	`
						", failed.  This is the error for the 32 bit "	+	`
						"attempt."
			ErrorFormatter-	-Error         $Error[1]		`
							-Description   $strErrL			`
							-strLogFileFqG $strLogFileFqG

			ExitProcessing- -strLogFileFqG $strLogFileFqG
		}#End If - second ?
	}#End If - first ?
	#Restore error action
	$ErrorActionPreference = $strEarHoldL
	Return ,$oConnectL
}#End MDB Connect ############################################################


<#############################################################################
.Synopsis
Creates and returns a recordset object based on supplied query.

.Description
A Connection object is required for the operation.  If one is not spedified
  in call to this function, the $oDbConnectionG is assumed to exist in the
  calling script.
#>
Function MdbSelectRecords-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Object[]" )]
Param(	[Parameter( Mandatory = $True ) ]		`
		[string]$Query						,	`
		[int]$OpenStatic     = 3			,	`
		[int]$LockOptimistic = 3			,	`
		$Connection          = $Null		,	`
		$blnLogInitializedG  = $Null		,	`
		$strLogFileFqG       = $Null			`
		)
	If( -NOT $Connection )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'oDbConnectionG'
		$Connection = $oDbConnectionG
	}#End If - connection null

	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$RecordSetL = New-Object -ComObject ADODB.Recordset
	$strMsgL    = "  $Query"
	Write-Verbose -Message $strMsgL
	$strMsgL | Out-File	-FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append

	Trap{ ErrorFormatter- -Error $Error[0] -strLogFileFqG $strLogFileFqG }
	$RecordSetL.Open(	$Query			,	`
						$Connection		,	`
						$OpenStatic		,	`
						$LockOptimistic		`
					)
	Return ,$RecordsetL	#Notice the comma preceding the variable.
}#End MDB Select Records #####################################################


<#############################################################################
.Synopsis
Accepts a query and updates records accordingly.

.Description
A Connection object is required for the operation.  If one is not spedified
  in call to this function, the $oDbConnectionG is assumed to exist in the
  calling script.
#>
Function MdbUpdateRecords-()
{#############################################################################
[CmdletBinding()]
Param(	[Parameter( Mandatory = $True ) ]		`
		[string]$TableName					,	`
		[Parameter( Mandatory = $True ) ]		`
		[string]$WhereClause				,	`
		[Parameter( Mandatory = $True ) ]		`
		$DataTable							,	`
		[int]$OpenStatic     = 3			,	`
		[int]$LockOptimistic = 3			,	`
		$Connection          = $Null		,	`
		$blnLogInitializedG  = $Null		,	`
		$strLogFileFqG       = $Null			`
		)
	If( -NOT $Connection )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'oDbConnectionG'
		$Connection = $oDbConnectionG
	}#End If - connection null
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$WhereClause = ( EndChar- -String $WhereClause -EndChar -String ';' )	+	`
					$WhereClause
	$arrKeysL    = $DataTable.Keys
	$strQueryL   =	'SELECT * '	+	`
					'FROM '		+	`
					$TableName	+	`
					' '			+	`
					$WhereClause
	Write-Verbose -Message $strQueryL
	$strMsgL | Out-File	-FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append

	$RecordSetL = New-Object -ComObject ADODB.Recordset
	$RecordSetL.Open(	$strQueryL		,	`
						$Connection		,	`
						$OpenStatic		,	`
						$LockOptimistic		`
					)
	While ( $RecordSetL.EOF -ne $True  )
	{	ForEach( $strKeyL IN $arrKeysL )
		{	$RecordSetL.Fields.Item( $strKeyL ).value = $DataTable.$strKeyL }
		$RecordSetL.Update()
		$RecordSetL.MoveNext()
	}#WEnd - EoF
	$RecordSetL.Close()
}#End MDB Update Records #####################################################


<#############################################################################
.Synopsis
Pads a specified string with a specified character on the left and/or right.

.Description

#Accepts an input string and length (mandatory)
#Pads string with specified character to specified length.
#Padded on Left/Right/Both, favoring Left and defaulting to Left.
#This version truncates input if longer than $Length
#>
Function PadString-()
{#############################################################################
[CmdletBinding()]
Param(	[Parameter( Mandatory = $True ) ] $StringToPad	,	`
		[Parameter( Mandatory = $True ) ] [int]$Length	,	`
		[string]$PadChar = " "							,	`
		[switch]$Left									,	`
		[switch]$Right									,	`
		[switch]$Both									,	`
		$strLogFileFqG   = $Null							`
		)
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - strLogFileFqG Null

	If( -NOT $StringToPad ) { $StringToPad = '' }
	$StringToPad = $StringToPad.ToString()
	$intDiffL = $Length - $StringToPad.Length
	If( $intDiffL -lt 0 )
	{	$strMsgL =	"`r`n'**Warning - Padded length less than "		+	`
					'input string length'							+	`
					"`r`n  Invalid call to PadString-"				+	`
					"`r`n    Specified padded length of output, "	+	`
					$Length											+	`
					".`r`n    Input string '"						+	`
					$StringToPad									+	`
					"`r`n    Original length,"						+	`
					$StringToPad.length								+	`
					".`r`n  Returning original string..."
		Write-Output $strMsgL -Foreground 'Yellow' -Background 'Black'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		Return $StringToPad
	}#End If - too short

	If( !$Left -AND !$Right -AND !$Both ){ $Left = $True }
	If( $Left )
	{	For(	$int1L = 1				;	`
				$int1L -le $intDiffL	;	`
				$int1L ++					`
			)
		{	$StringToPad = $PadChar + $StringToPad }
	}#End If - Left

	If( $Right )
	{	For(	$int1L = 1				;	`
				$int1L -le $intDiffL	;	`
				$int1L ++					`
			)
		{	$StringToPad += $PadChar }
	}#End If - Right

	If( $Both )
	{	$intLeftL  = [int]($intDiffL / 2 )
		$intRightL = $intDiffL - $intLeftL

		#Pad left side...
		For(	$int1L = 1				;	`
				$int1L -le $intLeftL	;	`
				$int1L ++					`
			)
		{	$StringToPad = $PadChar + $StringToPad }

		#Pad right side...
		For(	$int1L = 1				;	`
				$int1L -le $intRightL	;	`
				$int1L ++					`
			)
		{	$StringToPad += $PadChar }
	}#End If - Both
	Return $StringToPad
}#End Pad String #############################################################


<#############################################################################
.Synopsis
Returns supplied text in a paragraph format, based on supplied parameters.

.Description
Accepts a string of text and optional parameters to reformat text to
  specified paragraph/line dimensions.
LeftMargin is whitespace added to each line as spaces.
IndentFirst is an optional indent for the first line, added to the value
  of LeftMargin.
TrimFirst removes all left-padding from first line, giving text of same
  width as rest of lines, without left-padding.
Default return is a formatted string with embedded new line characters.
-ReturnArray switch causes return of an array of text lines.
#>
Function Paragraph-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.String" )]
Param(	[string]$Text					,	`
		[int]$RightMargin    = 78		,	`
		[int]$LeftMargin     = 4		,	`
		[int]$IndentFirst    = 0		,	`
		[switch]$TrimFirst   			,	`
		[switch]$ReturnArray			,	`
		$blnLogInitializedG  = $Null	,	`
		$strLogFileFqG       = $Null		`
		)
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$arrInputL        = $Text.Split( " " )
	$intLastIncludedL = 0
	$intTrimAtL       = 0
	$strExtraL        = ""
	$strLMarginPadL   = " " * $LeftMargin
	$strTextPieceL    = ""

	#Not sure why...
	If( $RightMargin - $LeftMargin -le 13 )
	{	$strMsgL = "  Minimum text width, LineWidth - Margin is 14 characters."
		Write-Output $strMsgL -Foreground 'Red'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		ExitProcessing- -strLogFileFqG $strLogFileFqG
		Break
	}#End If width 13

	If( $TrimFirst ){ $intTextWidthL = $RightMargin }
	Else            { $intTextWidthL = $RightMargin - $LeftMargin }

	#Add indent to first word/line
	$arrInputL[0] = ( " " * $IndentFirst ) + $arrInputL[0]

	If( $arrInputL[0].Length -gt $intTextWidthL )
	{	#First "word" is longer than first line.
		#Chop it off and assign the rest to $strExtraL
		$strTempL   = $arrInputL[ 0 ].SubString( 0, $intTextWidthL )
		$arrOutputL = @( $strTempL )
		$strExtraL  = $arrInputL[0].SubString( $intTextWidthL )
	}#End If - first word length
	Else	#First word fits first line
	{	$intLineCharCountL        = 0
		$strTempL                 = ""
		$intCurrentArrayPositionL = 0
		While( $intLineCharCountL -lt $intTextWidthL )
		{	For(	$int1L = $intCurrentArrayPositionL		;	`
					$int1L -le $arrInputL.GetUpperBound( 0 );	`
					$int1L ++ )
			{	If( $intLineCharCountL + $arrInputL[ $int1L ].Length	`
					-lt													`
					$intTextWidthL )
				{	#Current item will fit, add it to the line
					$strTempL += " " + $arrInputL[ $int1L ]

					#Update length of current line
					$intLineCharCountL += $arrInputL[ $int1L ].Length + 1

					#Update last item included
					$intLastIncludedL = $int1L

					#Test to see if just processed last input element
					If( $int1L -eq $arrInputL.GetUpperBound( 0 ) )
					{	$intLineCharCountL = $intTextWidthL }

				}#End If - Current item will fit
				Else	#TextWidth reached
				{	$intLineCharCountL = $intTextWidthL
					Break
				}#End Else
			}#Next int1L
		}#End While

		#Trim leading space
		$strTempL = $strTempL.Substring( 1 )
	}#End Else - First word fits first line

	If( $TrimFirst ){ $arrOutputL = @( $strTempL ) }
	Else{ $arrOutputL = $strLMarginPadL +  @( $strTempL ) }

	#First line complete with possible text in $strExtraL
	########################################################
	#Now process remaining lines, which will all be the same
	#  length, ( $RightMargin - $LeftMargin )
	$intTextWidthL = $RightMargin - $LeftMargin

	#Continue processing with next item, ( $intLastIncludedL + 1 )
	$blnContinueL = $True
	While( $blnContinueL )
	{	If( $strExtraL.Length -gt $intTextWidthL )
		{	#Left over "word" is longer than max line width
			#Trim line to length
			$strTempL = $strExtraL.Substring( 0, $intTextWidthL )

			#Add leading spaces
			$strTempL = $strLMarginPadL + $strTempL

			#Add to ouptut array
			$arrOutputL += $strTempL

			#Trim processed text from leftover string
			$strExtraL = $strExtraL.Substring( $intTextWidthL )

		}#End If - leftover too long
		Else	#Leftover not too long
		{	#Process $strExtraL and rest of input...
			If( $arrInputL.GetUpperBound( 0 ) -eq 0 )
				#Input is only one "word" but wraps
			{	$arrOutputL  += $strLMarginPadL + $strExtraL
				$blnContinueL = $False
			}#End If - Single long word

			$int1L = $intLastIncludedL + 1
			For(	$int1L = $int1L				 			;	`
					$int1L -le $arrInputL.GetUpperBound( 0 );	`
					$int1L ++ )
			{	#Is Extra plus next word greater than line width?
				If(	$arrInputL[ $int1L ].Length	+	`
					$strExtraL.Length			+	`
					1							`
					-gt								`
					$intTextWidthL )
				{	#Line composed of $strExtraL and current item is too long.
					#Check to see if $strExtraL + space is GE TextWidth
					If( $strExtraL.Length + 1 -ge $intTextWidthL )
					{	#Create an output line with only Extra
						$arrOutputL += $strLMarginPadL + $strExtraL

						#Trim $strExtraL to eliminate it
						$strExtraL = ""

						#Push pointer back.  Need to process same item
						$int1L --
					}#End If
					Else
					{	#Join both, with space.
						#Trim current item to fit line length.
						#Assign the rest of current item to $strExtraL.
						If( $strExtraL -eq "" ){ $strTempL = "" }
						Else                   { $strTempL = $strExtraL + " " }

						$intTrimAtL    = $intTextWidthL - $strTempL.Length
						$strTextPieceL = $arrInputL[ $int1L ].SubString(	`
															0, $intTrimAtL )
						$strTempL = $strLMarginPadL + $strTextPieceL
						$arrOutputL += $strTempL

						#Create the leftover
						$strExtraL = $arrInputL[ $int1L ].SubString( $intTrimAtL )

					}#End Else - Extra fills line without any of next word
				}#End If - Extra plus next word wraps
				Else
				{	#Line will hold at least $strExtraL and complete next word.
					#Append into holding string.
					$strTempL = $strExtraL + " " + $arrInputL[ $int1L ]
					$strTempL = $strTempL.Trim()

					#Trim leftover holder
					$strExtraL = ""

					#Set line position counter
					$intLineCharCountL = $strTempL.Length

					#Increment array pointer.
					$int1L ++

					While(	$intLineCharCountL -lt $intTextWidthL		`
							-AND										`
							$int1L -le $arrInputL.GetUpperBound( 0 )	`
							)
					{	For(	$int1L = $int1L							;	`
								$int1L -le $arrInputL.GetUpperBound( 0 );	`
								$int1L ++ )
						{	If( $intLineCharCountL			`
								+							`
								$arrInputL[ $int1L ].Length	`
								-lt							`
								$intTextWidthL				`
								)
							{	#Current item will fit, add it to the line
								$strTempL += " " + $arrInputL[ $int1L ]

								#Update length of current line
								$intLineCharCountL = $strTempL.Length

								#Update last item included
								$intLastIncludedL = $int1L
							}#End If
							Else
							{	$intLineCharCountL = $intTextWidthL
								$int1L = $intLastIncludedL
								Break
							}#End Else
						}#Next
					}#End While
					$arrOutputL += @( "`n$strLMarginPadL" + $strTempL )
				}#End Else
			}#Next - input element
				If( $int1L -ge $arrInputL.GetUpperBound( 0 ) )
					{ $blnContinueL = $False }
		}#End Else - Leftover not too long
	}#WEnd - Continue

	If( $ReturnArray )
		{ Return ,$arrOutputL }
	Else	#Return a string
		{
		$strReturnL = ""
		ForEach( $strItemL in $arrOutputL )
			{ $strReturnL += "$strItemL`n" }
		Return $strReturnL
		}#End Else
}#End Paragraph ##############################################################


<#############################################################################
.Synopsis
Reads specified file, typically for configuration items.

.Description
Reads file and returns either an array of text lines or a hashtable of
  key/value pairs.
The -Hashtable option removes comment/blank lines then splits remaining
  lines on the specified character, -Delimiter.
The first element from the line is trimmed and becomes the entry Key.
The remainining element(s) are not trimmed, and become the entry Value.
#>
Function ReadInputTextFile-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Object[]" )]
Param(	#Fully qualified path to target file								`
		[Parameter( Mandatory = $True ) ]									`
		[string]$InputFile												,	`
		#Comment character used to ignore lines in input file				`
		[string]$CommentCharacter = '#'									,	`
		#Default folder for file if only file name as -InputFile			`
		[string]$DefaultPath = $Null									,	`
		#Switch to return key/value pairs rather than text array			`
		[switch]$ReturnHashTable										,	`
		#Switch to allow continue if input file not Found					`
		[switch]$NoAbort												,	`
		#Delimiter used to create key/value pairs for hashtable option		`
		[string]$Delimiter        = ' '									,	`
		#Switch, remove whitespace from lines, if not hashtable				`
		[switch]$TrimLine												,	`
		$strLogFileFqG = $Null												`
		)
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'blnLogInitializedG','strLogFileFqG'
	}#End If - null

	$colInputL = @()
	$strTargetFileL = $InputFile

	If( -NOT( Test-Path -Path $strTargetFileL ) )
	{	#Not found.  Check script directory.
		GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strScriptPathG'

		$strTempPathL   = EndChar- -String $strScriptPathG -EndChar '\'
		$strTargetFileL = $strTempPathL + $strTargetFileL

		If( -NOT( Test-Path -Path $strTargetFileL ) )
		{	If( $NoAbort )
			{	If( $ReturnHashTable )
				{	$htReturnL = @{ Result = 'FileNotFound' }
					Return $htReturnL
				}#End If
				Else
				{	$colInputL = @( 'FileNotFound' )
					Return ,$colInputL
				}#End Else
			}#End If - NoAbort

			$strMsgL =	'**Error - Specified input file not found: '	+	`
						"`r`n`t$InputFile"								+	`
						"`r`n  Also not found as:"						+	`
						"`r`n`t$strTargetFileL"							+	`
						"`r`n  Exiting..."
			Write-Output $strMsgL -Foreground 'Red'
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
			ExitProcessing- -strLogFileFqG $strLogFileFqG
		}#End If - file not found in script folder either
	}#End If - file not found

	$arrInputL = Get-Content -Path $strTargetFileL
	$colTargetsL = @()

	ForEach( $strLineL IN $arrInputL )
	{	$strTempL = $strLineL.Trim()
		If( -NOT( $strTempL.StartsWith( $CommentCharacter ) ) )
		{	If( -NOT( $strTempL -eq '' ) )
			{	If( $TrimLine ){ $strTempL = $strTemp.Trim() }
				$colTargetsL += ,$strTempL }
		}#End If - #
	}#Next - strLineL

	If( $ReturnHashTable )
	{	$htReturnL = @{}
		ForEach( $strLineL IN $colTargetsL )
		{	#Break the line, on -Delimiter.
			$arrTempL = $strLineL.Split( $Delimiter )

			#Assign first element as key for hashtable entry.
			$strKeyL  = ( [string]$arrTempL[0] ).Trim()

			#In case there were multiple -Delimiters in the line,
			#  rebuild the line, reinserting delimiters.
			$strTempL = ''
			For(	$int1L = 1								;	`
					$int1L -le $arrTempL.GetUpperBound(0)	;	`
					$int1L ++									`
				)
			{	$strTempL += $arrTempL[ $int1L ] + $Delimiter }

			#Remove trailining Delimiter
			$strTempL  = $strTempL.Substring( 0, ( $strTempL.Length - 1 ) )
			$strValueL = [string]$strTempL
			$htReturnL.$strKeyL = $strValueL
		}#Next - strLineL
		Return $htReturnL
	}#End if - ReturnHashTable
	Else { 	Return ,$colTargetsL }
}#End Read Input Text File ###################################################


<#############################################################################
.Synopsis
Sends SMTP message using specified parameters.

.$DESCRIPTION
If -SmtpServer argument is missing, assume use of all values of the
  Script-level oSmtpServerG object.
#>
Function SendEmailPS2-()
{#############################################################################
Param(	$BccTo              = $Null	,	`
		$CcTo               = $Null	,	`
		$MsgBody            = $Null	,	`
		$UserFrom           = $Null	,	`
		$Recipient          = $Null	,	`
		$SmtpServer         = $Null	,	`
		$SmtpPort           = $Null	,	`
		$Subject            = $Null	,	`
		$UseCred            = $Null	,	`
		$UserAuth           = $Null	,	`
		$Password           = $Null	,	`
		[string]$Attachments		,	`
		[switch]$HTML				,	`
		[switch]$UseSsl				,	`
		$SmtpConfig					,	`
		$LogFile            = $Null	,	`
		$Verbose            = $False	`
		)
	$blnNoSendL = $False
	$strErrorL  = ''

	#If incoming parms are null/missing, substitute $oSmtpServerG values
	If( -NOT $BccTo   ){ $BccTo   = '' }
	If( -NOT $CcTo    ){ $CcTo    = '' }
	If( -NOT $MsgBody ){ $MsgBody = '' }

	If( -NOT $UserFrom )
	{	$UserFrom = $SmtpConfig.EmailFrom
		If( -NOT( $UserFrom.Contains( '@' ) ) )
		{	$strErrorL += "`r`n  Message must have a sender address."
			$blnNoSendL = $True
		}#End If - no @
	}#End If - UserFrom

	If( -NOT $Recipient )
	{	$Recipient = $SmtpConfig.EmailTo
		If( -NOT( $Recipient.Contains( '@' ) ) )
		{	$strErrorL += "`r`n  Message must have a recipient."
			$blnNoSendL = $True
		}#End If - no @
	}#End If - Recipient

	If( -NOT $SmtpServer )
	{	$SmtpServer = $SmtpConfig.ServerName
		If( -NOT( $SmtpServer.Contains( '.' ) ) )
		{	$strErrorL += "`r`n  Message must have an SMTP server."
			$blnNoSendL = $True
		}#End If - no .
	}#End If - SmtpServer

	If( -NOT $SmtpPort )
	{	$SmtpPort = [int]$SmtpConfig.Port
		If( $SmtpPort -eq 0 ){	$SmtpPort = 25 }
	}#End If - port null

	If( -NOT $Subject )
	{	$Subject = $SmtpConfig.Subject
		If( [string]$Subject -eq '')
		{	$strMsgL = "  Message should have a subject."
			Write-Output $strMsgL -Foreground 'Yellow' -Background 'Black'
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
			$strErrorL += "`r`n  Message must have a subject."
		}#End If - Subject2 ''
	}#End If - Subject null

	If( -NOT $UseCred ){ $UseCred = $SmtpConfig.UseAuth }
	If( $UseCred -eq $True )
	{	If( -NOT $UserAuth )
		{	$UserAuth = $SmtpConfig.AuthName
			If( -NOT( $UserAuth.Contains( '@' ) ) )
			{	$strErrorL +=	"`r`n  Message must have an authorized "	+	`
								'user name if SMTP Auth is used.'
				$blnNoSendL = $True
			}#End If - no @
		}#End If - UserAuth Null

		If( -NOT $Password )
		{	$Password = $SmtpConfig.AuthPwd
			If( [string]$Password.Length -lt 1 )
			{	$strErrorL +=	"`r`n  Message must have a password "	+	`
								'if SMTP Auth is used.'
				$blnNoSendL = $True
			}#End If - no password
		}#End If - Password Null
	}#End If - UseCred True

	If( $blnNoSendL )
	{	$strMsgL =	'**Error - email message cannot be sent'			+	`
					"`r`n  The following condition(s) prevent this "	+	`
					'message from being sent:'							+	`
					$strErrorL
		Write-Output $strMsgL -Foreground 'Red'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		Return $False
	}#End If - no send

	$blnReturnL   = $False
	$htParmsL     = @{}
	$strErrorL    = ''

	$htParmsL.SmtpServer = $SmtpServer
	$htParmsL.From       = $UserFrom
	$htParmsL.To         = $Recipient
	$htParmsL.Subject    = $Subject
	$htParmsL.Body       = $MsgBody

	If( $BccTo -ne '' ) { $htParmsL.BCC    = $BccTo  }
	If( $CcTo -ne '' )  { $htParmsL.CC     = $CcTo   }
	If( $HTML )         { $htParmsL.HTML   = $True   }
	If( $UseSsl )       { $htParmsL.UseSsl = $True }

	If( $Attachments )
	{	#This can be a comma-delimited string or array.
		If( $Attachments -IS [Array] ){ $htParmsL.Attachments =	$Attachments }
		Else
		{	$arrTempL = $Attachments.Split( ',' )
			$htParmsL.Attachments =	$arrTempL
		}#End Else - not array
	}#End If - Attachments

	If( $UseCred )
	{	$strMsgL = '  Specifying user credentials...'
		If( $UserAuth ){ $strMsgL += "`r`n    User: $UserAuth" }
		Else
		{	$strErrorL +=	"`r`n  **Error - Authorized user not listed."	+	`
							"`r`n      Required when using credentials."
		}#End Else - no user

		If( $Password ){ $strMsgL += "`r`n     Pwd: $Password" }
		Else
		{	$strErrorL +=	"`r`n  **Error - Password not listed."			+	`
							"`r`n      Required when using credentials."
		}#End Else - no Password

		#Make Credential and add to HT
		#$htParmsL.Credential = $oCredentialL
	}#End If - UseCred
	Else
	{	$strMsgL = '  Sending with anonymous relay...' }
	Write-Verbose -Message $strMsgL

	If( $strErrorL -ne '' )
	{	$strMsgL =	'  **The following problem(s) prevent '	+	`
					'sending this message.'					+	`
					"`r`n    $strErrorL"
		Write-Output $strMsgL -Foreground 'Red'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
	}#End If - Error
	Else
	{	$htParmsL.EA = 'Continue'
		If( $Verbose )
		{	$strParmsL = '  Parameters for mail message...'
			$arrKeysL = $htParmsL.Keys
			ForEach( $strKeyL IN $arrKeysL )
			{	$strParmsL +=	"`r`n`t"			+	`
								$strKeyL			+	`
								": "				+	`
								$htParmsL.$strKeyL
			}#Next - key
			Write-Output $strMsgL
			$strMsgL | Out-File	-FilePath $LogFile	`
								-Encoding 'ASCII'	`
								-Append
		}#End If - Verbose

		$strMsgL = '  Sending...'
		Write-Verbose -Message $strMsgL
		$strMsgL | Out-File	-FilePath $LogFile	`
							-Encoding 'ASCII'	`
							-Append
		Send-MailMessage @htParmsL

		If( $? )
		{	$blnReturnL = $True
			$strMsgL = '    Completed successfully.'
			Write-Verbose -Message $strMsgL
			$strMsgL | Out-File	-FilePath $LogFile	`
								-Encoding 'ASCII'	`
								-Append
		}#End If
		Else
		{	ErrorFormatter-	-Error         $Error[0]		`
							-Description   $strErrL			`
							-strLogFileFqG $strLogFileFqG
		}#End Else
	}#End Else - no error
	Return $blnReturnL
}#End Send Email PS 2 ########################################################


<#############################################################################
.Synopsis
Sends SMTP message using specified parameters.

.$DESCRIPTION
If arguments available in $oSmtpServerG are missing, values from that object
  are used.
#>
Function SendEmailPS-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Boolean" )]
Param(	[string]$BccTo				,	`
		[string]$CcTo				,	`
		[string]$MsgBody			,	`
		[string]$UserFrom			,	`
		[string]$Recipient			,	`
		[string]$SmtpServer 		,	`
		$SmtpPort           = $Null	,	`
		[string]$Subject			,	`
		$UseCred            = $Null	,	`
		[string]$UserAuth        	,	`
		[string]$Password        	,	`
		[string]$Attachments		,	`
		[switch]$HTML				,	`
		[switch]$UseSsl				,	`
		$strLogFileFqG      = $Null	,	`
		$oSmtpServerG					`
		)
	If( -NOT $oSmtpServerG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'oSmtpServerG,Verbose2'
	}#End If
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$blnNoSendL = $False
	$strErrorL  = ''

	If( $UserFrom -eq '' )
	{	$UserFrom = $oSmtpServerG.EmailFrom
		If( -NOT( $UserFrom.Contains( '@' ) ) )
		{	$strErrorL += "`r`n  Message must have a sender address."
			$blnNoSendL = $True
		}#End If - no @
	}#End If - UserFrom

	If( $Recipient -eq '' )
	{	$Recipient = $oSmtpServerG.EmailTo
		If( -NOT( $Recipient.Contains( '@' ) ) )
		{	$strErrorL += "`r`n  Message must have a recipient."
			$blnNoSendL = $True
		}#End If - no recipient
	}#End If - Recipient

	If( $SmtpServer -eq '' )
	{	$SmtpServer = $oSmtpServerG.ServerName
		If( -NOT( $SmtpServer.Contains( '.' ) ) )
		{	$strErrorL += "`r`n  Message must have an SMTP server."
			$blnNoSendL = $True
		}#End If - no .
	}#End If - SmtpServer

	If( -NOT $SmtpPort )
	{	$SmtpPort = [int]$oSmtpServerG.Port
		If( $SmtpPort -eq 0 ){	$SmtpPort = 25 }
	}#End If - port null

	If( $Subject -eq '' )
	{	$Subject = [string]$oSmtpServerG.Subject
		If( $SmtpServer -eq '')
		{	$strMsgL = "  Message should have a subject."
			Write-Output $strMsgL -Foreground 'Yellow' -Background 'Black'
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
			$strErrorL += "`r`n  Message must have an SMTP server."
		}#End If - server ''
	}#End If - SmtpServer null

	If( -NOT $UseCred ){ $UseCred = $oSmtpServerG.UseAuth }
	If( $UseCred -eq $True )
	{	If( $UserAuth -eq '' )
		{	$UserAuth = $oSmtpServerG.AuthName
			If( -NOT( $UserAuth.Contains( '@' ) ) )
			{	$strErrorL +=	"`r`n  Message must have an authorized "	+	`
								'user name if SMTP Auth is used.'
				$blnNoSendL = $True
			}#End If - no @
		}#End If - UserAuth Null

		If( $Password -eq '' )
		{	$Password = [string]$oSmtpServerG.AuthPwd
			If( $Password -eq '' )
			{	$strErrorL +=	"`r`n  Message must have a password "	+	`
								'if SMTP Auth is used.'
				$blnNoSendL = $True
			}#End If - no password
		}#End If - Password Null
	}#End If - UseCred True

	If( $blnNoSendL )
	{	$strMsgL =	'**Error - email message cannot be sent'			+	`
					"`r`n  The following condition(s) prevent this "	+	`
					'message from being sent:'							+	`
					$strErrorL
		Write-Output $strMsgL -Foreground 'Red'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		Return $False
	}#End If - no send

	$blnReturnL   = $False
	$htParmsL     = @{}
	$strErrorL    = ''

	$htParmsL.SmtpServer = $SmtpServer
	$htParmsL.From       = $UserFrom
	$htParmsL.To         = $Recipient
	$htParmsL.Subject    = $Subject
	$htParmsL.Body       = $MsgBody

	If( $BccTo -ne '' ) { $htParmsL.BCC    = $BccTo  }
	If( $CcTo -ne '' )  { $htParmsL.CC     = $CcTo   }
	If( $HTML )         { $htParmsL.HTML   = $True   }
	If( $UseSsl )       { $htParmsL.UseSsl = $UseSsl }

	If( $Attachments -ne '' )
	{	#This can be a comma-delimited string or array.
		If( $Attachments -IS [Array] ){ $htParmsL.Attachments =	$Attachments }
		Else
		{	$arrTempL = $Attachments.Split( ',' )
			$htParmsL.Attachments =	$arrTempL
		}#End Else - not array
	}#End If - Attachments

	If( $UseCred )
	{	$strMsgL = '  Specifying user credentials...'
		If( $UserAuth ){ $strMsgL += "`r`n    User: $UserAuth" }
		Else
		{	$strErrorL +=	"`r`n  **Error - Authorized user not listed."	+	`
							"`r`n      Required when using credentials."
		}#End Else - no user

		If( $Password ){ $strMsgL += "`r`n     Pwd: $Password" }
		Else
		{	$strErrorL +=	"`r`n  **Error - Password not listed."			+	`
							"`r`n      Required when using credentials."
		}#End Else - no Password

		#Make Credential and add to HT
		#$htParmsL.Credential = $oCredentialL
	}#End If - UseCred
	Else
	{	$strMsgL = '  Sending with anonymous relay...' }
	Write-Verbose -Message $strMsgL

	If( $strErrorL -ne '' )
	{	$strMsgL =	'  **The following problem(s) prevent '	+	`
					'sending this message.'					+	`
					"`r`n    $strErrorL"
		Write-Output $strMsgL -Foreground 'Red'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
	}#End If - Error
	Else
	{	$htParmsL.EA = 'Continue'
		If( $Verbose2 )
		{	$strParmsL = '  Parameters for mail message...'
			$arrKeysL = $htParmsL.Keys
			ForEach( $strKeyL IN $arrKeysL )
			{	$strParmsL +=	"`r`n`t"			+	`
								$strKeyL			+	`
								": "				+	`
								$htParmsL.$strKeyL
			}#Next - key
			Write-Output $strMsgL
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
		}#End If - Verbose

		$strMsgL = '  Sending...'
		Write-Verbose -Message $strMsgL
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		Send-MailMessage @htParmsL

		If( $? )
		{	$blnReturnL = $True
			$strMsgL = '    Completed successfully.'
			Write-Verbose -Message $strMsgL
			$strMsgL | Out-File	-FilePath $strLogFileFqG	`
								-Encoding 'ASCII'			`
								-Append
		}#End If
		Else
		{	$strErrL = 'Problem sending email.'
			ErrorFormatter-	-Error         $Error[0]		`
							-Description   $strErrL			`
							-strLogFileFqG $strLogFileFqG
		}#End Else
	}#End Else - no error
	Return $blnReturnL
}#End Send Email PS 1 ########################################################


<#############################################################################
.Synopsis
Converts a dateTime object to the format used by SQL Server.

.Description
Accepts a dateTime object, or anything that can be converted to a dateTime
  object and formats it to a SQL representation with optional "trimming"
  of precision.
Option to add single quotes surrounding text.
#>
Function SqlDateTime-()
{#############################################################################
Param(	$DateTime = (Get-Date)	,	`
		[switch]$TrimHours		,	`
		[switch]$TrimMinutes	,	`
		[switch]$TrimSeconds	,	`
		[switch]$TrimMillis		,	`
		[switch]$AddQuotes			`
		)
	$DateTime = [dateTime]$DateTime

	#Create date as string in format: CCYY-MM-DD HH:MM:SS.hhh
	If( $TrimHours )
	{	$TrimMinutes = $True
		$TrimSeconds = $True
		$TrimMillis  = $True
	}#End If - Hours
	Else
	{	If( $TrimMinutes )
		{	$TrimSeconds = $True
			$TrimMillis  = $True
		}#End If - Minutes
		Else
		{	If( $TrimSeconds ){ $TrimMillis = $True } }
	}#End Else - Not hours

	$strDateTimeL = [string]( $DateTime.Year )

	$strTempL = [string]( $DateTime.Month )
	If( $strTempL.Length -eq 1 ){ $strTempL = '0' + $strTempL }
	$strDateTimeL += "-$strTempL"

	$strTempL = [string]( $DateTime.Day )
	If( $strTempL.Length -eq 1 ){ $strTempL = '0' + $strTempL }
	$strDateTimeL += "-$strTempL"

	If( $TrimHours )
	{	$strDateTimeL += ' 00:00:00.000' }
	Else
	{	$strTempL = [string]( $DateTime.Hour )
		If( $strTempL.Length -eq 1 ){ $strTempL = '0' + $strTempL }
		$strDateTimeL += " $strTempL"

		If( $TrimMinutes )
		{	$strDateTimeL += ':00:00.000' }
		Else
		{	$strTempL = [string]( $DateTime.Minute )
			If( $strTempL.Length -eq 1 ){ $strTempL = '0' + $strTempL }
			$strDateTimeL += ":$strTempL"

			If( $TrimSeconds )
			{	$strDateTimeL += ':00.000' }
			Else
			{	$strTempL = [string]( $DateTime.Second )
				If( $strTempL.Length -eq 1 ){ $strTempL = '0' + $strTempL }
				$strDateTimeL += ":$strTempL"

				If( $TrimMillis )
				{	$strDateTimeL += '.000' }
				Else
				{	$strTempL = [string]( $DateTime.Millisecond )
					While( $strTempL.Length -lt 3 )
					{	$strTempL = '0' + $strTempL }
					$strDateTimeL += ".$strTempL"
				}#End Else - Millisecond
			}#End Else - Seconds
		}#End Else - Minutes
	}#End Else - Hours

	If( $AddQuotes ){ $strDateTimeL = "'" + $strDateTimeL + "'" }
	Return $strDateTimeL
}#End SQL Date Time ##########################################################


<#############################################################################
.Synopsis
Process a string to be included in a SQL query, adding/escaping any
  single quotes to prevent unwanted "breaks" in query processing.

.Description
For instance,    SELECT * FROM TableName WHERE LName = "O'Leary"
  is changed to, SELECT * FROM TableName WHERE LName = "O''Leary"
#>
Function SqlEscapeQuotesInString-( [string]$InputString )
{#############################################################################
	$strOutputL = $InputString.Replace( "'", "''" )
	Return $strOutputL
}#End SQL Escape Quotes In String ############################################


<#############################################################################
.Synopsis
Executes any fully formed query.

.Description
Uses supplied or calling script default connection object.
#>
Function SqlExecuteQuery-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Void" )]
Param(	[ Parameter( Mandatory = $True ) ] [string]$Query	,	`
		$Connection    = $Null								,	`
		$strLogFileFqG = $Null									`
		)
	If( -NOT $Connection )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'oDbConnectionG'
		$Connection = $oDbConnectionG
	}#End If - connection null
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$strMsgL = "  Executing Query: $Query"
	Write-Verbose -Message -Message $strMsgL
	$strMsgL | Out-File	-FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append

	$oSqlCommandL = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandText = $Query
	$oSqlCommandL.Connection  = $Connection

	[Void]$oSqlCommandL.ExecuteNonQuery()
	[Void]$oSqlCommandL.Dispose()
}#End Function - SQL Execute Query ###########################################


##############################################################################
Function SqlInitializeDataReader-()
{#############################################################################
[CmdletBinding()]
Param(	[string]$Query			,	`
		$Connection    = $Null	,	`
		$strLogFileFqG = $Null		`
		)
	If( -NOT $Connection )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'oDbConnectionG'
		$Connection = $oDbConnectionG
	}#End If - connection null

	If( ( -NOT $strLogFileFqG ) )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$strMsgL = "  Initializing dataReader with query: $Query"
	Write-Verbose -Message $strMsgL
	$strMsgL | Out-File	-FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append

	$oSqlCommandL = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandTimeout = 30
	$oSqlCommandL.Connection     = $Connection
	$oSqlCommandL.CommandText    = $Query

	If( -NOT $? )
	{	$strErrL = "While creating new SQL Command object."
		ErrorFormatter-	-Error         $Error[0]		`
						-Description   $strErrL			`
						-strLogFileFqG $strLogFileFqG
	}#End If - Error

	$oSqlReaderL = $oSqlCommandL.ExecuteReader()
	If( -NOT $? )
	{	$strErrL = "While executing SQL Reader with, $Query"
		ErrorFormatter-	-Error         $Error[0]		`
						-Description   $strErrL			`
						-strLogFileFqG $strLogFileFqG
	}#End If - Error
	Return ,$oSqlReaderL	#Notice comma preceding variable
}#End Function - SQL Initialize Reader #######################################


<#############################################################################
.Synopsis
Creates and returns a SQL recordset based on supplied query.

.Description
Uses supplied or calling script default connection object.
Note: This function uses GetCallerVariable- to retrieve the Connection
        object.  This means if this function is not called directly from
        the primary script, the connection object must be passed in.
#>
Function SqlInitializeDataset-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Object[]" )]
Param(	[string]$Query				,	`
		$Connection         = $Null	,	`
		$strLogFileFqG      = $Null		`
		)
	If( -NOT $Connection )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'oDbConnectionG'
		$Connection = $oDbConnectionG
	}#End If - connection null
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'

	$oSqlCommandL = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandText    = $Query
	$oSqlCommandL.Connection     = $Connection
	$oDataAdapterL               = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter
	$oDataAdapterL.SelectCommand = $oSqlCommandL
	}#End If - null

	$strMsgL = "  Executing query, $Query"
	Write-Verbose -Message $strMsgL

	# Create and fill the DataSet object
	$oDataSetL = New-Object -TypeName System.Data.DataSet
	$oDataAdapterL.Fill( $oDataSetL ) | Out-Null
	Return ,$oDataSetL
}#End SQL Initialize Dataset #################################################


##############################################################################
Function SqlInitializeDbConnection-()
{#############################################################################
[CmdletBinding()]
Param(	[string]$Server					,	`
		[string]$Database				,	`
		[string]$Port       = "1433"	,	`
		[string]$Instance   = ""		,	`
		[string]$UserId     = "Windows"	,	`
		[string]$Pwd        = ""		,	`
		$strLogFileFqG      = $Null		,	`
		[bool]$Verbose2     = $False		`
		)
	If( $Verbose2 ){ $VerbosePreference = 'Continue' }
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - null

	#Base connection string
	$strConnectionL =	'Data Source='	+	`
						$Server			+	`
						','				+	`
						$Port

	#Add instance name if needed
	If( $Instance -ne "" ){ $strConnectionL += "\$Instance" }

	#Add database name
	$strConnectionL += "; Initial Catalog=$Database; "

	#Determine auth mode
	If ( $UserId.ToUpper() -eq "WINDOWS" )
	{	$strConnectionL += "Integrated Security = SSPI;" }
	Else #SQL standard security
	{	$strConnectionL +=	"User ID='$UserId'; Password='$Pwd'; "	+	`
							"Trusted_Connection=No;"
	}#End Else

	$strMsgL = "  Executing DB Connection with string: $strConnectionL"
	Write-Verbose -Message $strMsgL
	$strMsgL | Out-File	-FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append

	$oDbConnectionL = New-Object -TypeName System.Data.SqlClient.SqlConnection
	$oDbConnectionL.ConnectionString = $strConnectionL
	If( -NOT $? )
	{	$strErrL = "While establishing connection string: $strConnectionL"
		ErrorFormatter-	-Error         $Error[0]		`
						-Description   $strErrL			`
						-strLogFileFqG $strLogFileFqG
	}#End If - Error

	$oDbConnectionL.Open()
	If( -NOT $? )
	{	$strErrL =	"While opening connection with "	+	`
					"connection string: $strConnectionL"
		ErrorFormatter-	-Error         $Error[0]		`
						-Description   $strErrL			`
						-strLogFileFqG $strLogFileFqG
	}#End If - Error
	Return $oDbConnectionL
}#End SQL Initialize DB Connection ###########################################


<#############################################################################
.Synopsis
Inserts a new record in specified table.

.Description
Uses supplied or calling script default connection object.
Accepts TableName and hashtable.
Hash table is column name (key) and data (value) for one record.
Builds lists of column names and values for query.
Adds .Parameters.AddWithValue, for each column, to command object, then
  executes command.
Records to be added must meet any required constraints, such as unique
  key values.
SqlUpdateRecords-() can also be used to add/insert records.  It checks for
  existence and updates or adds if not found.
#>
Function SqlInsertRecord-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Void" )]
Param(	[string]$TableName			,	`
		$DataTable					,	`
		$Connection         = $Null	,	`
		$blnLogInitializedG = $Null	,	`
		$strLogFileFqG      = $Null		`
		)
	If( -NOT $Connection )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'oDbConnectionG'
		$Connection = $oDbConnectionG
	}#End If - connection null
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$strFieldL               = ''
	$strParmNameL            = ''
	$strTableL               = ''
	$arrFieldNamesL          = $DataTable.Keys
	$oSqlCommandL            = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$oSqlCommandL.Connection = $Connection

	#Create column name list.
	$strFieldsListL = ''
	ForEach( $strColumnNameL IN $arrFieldNamesL )
	{	$strFieldsListL += "$strColumnNameL, " }

	#Trim trailing ", "
	$strFieldsListL = $strFieldsListL.Substring( 0, $strFieldsListL.Length - 2 )

	#Create data values list.
	$strValuesListL = ''
	ForEach( $strColumnNameL IN $arrFieldNamesL )
		{	$oColumnValueL = $DataTable.$strColumnNameL
			#Add to values list
			$strValuesListL += "$oColumnValueL, "

			#Add parm to command object
			$strParmL = "@$strColumnNameL"
			[void]$oSqlCommandL.Parameters.AddWithValue(	$strParmL		,	`
															$oColumnValueL		`
														)
		}#Next
	#Trim trailing ", " from values list
	$strValuesListL = $strValuesListL.Substring( 0, $strValuesListL.Length - 2 )

	#Build query string
	$strQueryL =	'INSERT INTO '	+	`
					$TableName		+	`
					'( '			+	`
					$strFieldsListL	+	`
					' ) VALUES( '	+	`
					$strValuesListL	+	`
					' )'
	$strMsgL = "  Insert Query: $strQueryL"
	WriteVerbose- -InputText $strMsgL

	$oSqlCommandL.CommandText = $strQueryL
	$intRowsL = $oSqlCommandL.ExecuteNonQuery()
	[void]$oSqlCommandL.Dispose()

	If( $intRowsL -ne 1 )
	{	$strMsgL =	"  **Error - While inserting record in table, $TableName."
		Write-Host $strMsgL -Foreground 'Red'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
		ExitProcessing- -strLogFileFqG $strLogFileFqG
	}#End If - Error
}#End SQL Insert Record ######################################################


<#############################################################################
.Synopsis
Checks to see if specified record exists.

.Description
Used by SqlUpdateRecords-() to verify before inserting.
#>
Function SqlRecordExists-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Void" )]
Param(	[string]$TableName			,	`
		[string]$WhereClause		,	`
		$Connection         = $Null	,	`
		$strLogFileFqG      = $Null		`
		)
	If( -NOT $Connection )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'oDbConnectionG'
		$Connection = $oDbConnectionG
	}#End If - connection null
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'strLogFileFqG'
	}#End If - null

	$strQueryL =	'SELECT * FROM '	+	`
					$TableName				+	`
					' '					+	`
					$WhereClause
	Write-Verbose -Message $strQueryL
	$strMsgL | Out-File	-FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append

	$oSqlCommandL = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandTimeout = 30
	$oSqlCommandL.Connection     = $Connection
	$oSqlCommandL.CommandText    = $strQueryL

	If( -NOT $? )
	{	$strErrL = "While creating new SQL Command object."
		ErrorFormatter-	-Error         $Error[0]		`
						-Description   $strErrL			`
						-strLogFileFqG $strLogFileFqG
	}#End If - Error

	$oSqlReaderL = $oSqlCommandL.ExecuteReader()
	If( -NOT $? )
	{	$strErrL = "While executing SQL Reader with, $strQueryL"
		ErrorFormatter-	-Error         $Error[0]		`
						-Description   $strErrL			`
						-strLogFileFqG $strLogFileFqG
	}#End If - Error

	$blnFoundL = $oSqlReaderL.HasRows
	[void]$oSqlReaderL.Close()
	Return $blnFoundL
}#End SQL Record Exists  #####################################################


<#############################################################################
.Synopsis
Removes all records from specified table.

.Description
Uses specified connection or default from calling script.
#>
Function SqlTruncateTable-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Void" )]
Param(	[string]$TableName	,	`
		$Connection = $Null		`
		)
	If( -NOT $Connection )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     '$oDbConnectionG'
		$Connection = $oDbConnectionG
	}#End If - connection null

	$strQueryL = "DELETE $TableName"
	$oSqlCommandL = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandText = $strQueryL
	$oSqlCommandL.Connection  = $Connection

	[Void]$oSqlCommandL.ExecuteNonQuery()
	[Void]$oSqlCommandL.Dispose()
}#End SQL Truncate Table #####################################################


<#############################################################################
.Synopsis
Updates an existing record or creates a new record with specified values.

.Description
Accepts TableName, Where clause and hashtable
DataTable is hashtable with column names as keys (name) and data values,
  for the fields whose values will be updated.

-Test confirms existence of record before attempting update.
-Create inserts record if -Test doesn't find it.

.Example
SqlUpdateRecords_	-TableName   'TableName'	`
					-WhereClause $strQueryL		`
					-DataTable   $htDataL		`
					-Test						`
					-Create
#>
Function SqlUpdateRecords-()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.Void" )]
Param(	[Parameter( Mandatory = $True, HelpMessage = 'TableName required' ) ]	`
		[string]$TableName													,	`
		[Parameter( Mandatory = $True, HelpMessage = 'WHERE required' ) ]		`
		[string]$WhereClause												,	`
		[Parameter( Mandatory = $True, HelpMessage = 'HT required' ) ]			`
		$DataTable															,	`
		[Switch]$Test														,	`
		[Switch]$Create														,	`
		$Connection         = $Null											,	`
		$strLogFileFqG      = $Null												`
		)
	If( -NOT $Connection )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet							`
							-SessionState $ExecutionContext.SessionState	`
							-VarNames     'oDbConnectionG'
		$Connection = $oDbConnectionG
	}#End If - connection null
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null

	If( $Test )
	{	If( -NOT( SqlRecordExists-	-Table         $TableName		`
									-WhereClause   $WhereClause		`
									-Connection    $Connection		`
									-strLogFileFqG $strLogFileFqG	`
				)
			)
		{	If( $Create )
			{	$strMsgL =	'  Record requested for Update does not exist.'	+	`
							"`r`n"											+	`
							'  Create option specified.'					+	`
							"`r`n"											+	`
							'  Record will be inserted.'
				Write-Verbose -Message $strMsgL
				$strMsgL | Out-File	-FilePath $strLogFileFqG	`
									-Encoding 'ASCII'			`
									-Append

				SqlInsertRecord-	-TableName     $TableName		`
									-DataTable     $DataTable		`
									-Connection    $Connection		`
									-strLogFileFqG $strLogFileFqG
				#Record created/inserted.  No update required.  Bail out.
				Return
			}#End If - Create
			Else
			{	$strMsgL =	'**Error - SQL record update failed.'		+	`
							"`r`n"										+	`
							'  Expected record, '						+	`
							$WhereClause								+	`
							"`r`n"										+	`
							'  In table, '								+	`
							$TableName									+	`
							"`r`n"										+	`
							'    was not found, and option to create '	+	`
							'new record was not selected.'
				Write-Output $strMsgL -Foreground 'Red'
				$strMsgL | Out-File	-FilePath $strLogFileFqG	`
									-Encoding 'ASCII'			`
									-Append
				ExitProcessing- -strLogFileFqG $strLogFileFqG
			}#End Else - no create
		}#End If - record doesn't exist
	}#End If - Test

	$strFieldL      = ""
	$strFieldsListL = ""
	$strParmNameL   = ""
	$strTableL      = ""

	#Create array of field/column names.
	$arrFieldNamesL = $DataTable.Keys

	#Make a string list of field names.
	ForEach( $strColumnNameL IN $arrFieldNamesL )
		{	$strFieldsListL += "$strColumnNameL, " }
	#Trim trailing characters.
	$strFieldsListL = $strFieldsListL.Substring( 0, $strFieldsListL.Length - 2 )

	#Build query string
	$strQueryL =	"UPDATE "		+	`
					$TableName		+	`
					" SET "
	ForEach( $strColumnNameL in $arrFieldNamesL )
	{	$strQueryL +=	$strColumnNameL				+	`
						"="							+	`
						$DataTable.$strColumnNameL	+	`
						","
	}#Next - strColumnNameL
	#Trim final comma
	$strQueryL = $strQueryL.Substring( 0, $strQueryL.Length - 1 )

	$strQueryL += " " + $WhereClause
	Write-Verbose -Message $strQueryL
	$strMsgL | Out-File	-FilePath $strLogFileFqG	`
						-Encoding 'ASCII'			`
						-Append

	$oSqlCommandL = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandText = $strQueryL
	$oSqlCommandL.Connection  = $Connection

	$intRowsL = $oSqlCommandL.ExecuteNonQuery()
	If( $intRowsL -ne 1 )
	{	$strMsgL =	"  **Error - While updating record(s) "	+	`
					"in table, "							+	`
					$strTableL
		Write-Output $strMsgL -Foreground 'Red'
		$strMsgL | Out-File	-FilePath $strLogFileFqG	`
							-Encoding 'ASCII'			`
							-Append
	}#End If - <> 1
	[Void]$oSqlCommandL.Dispose
}#End SQL Update Records #####################################################


<#############################################################################
.Synopsis
Pause of specified length with progress display.

.Description
Creates a pause and displays progress with Write-Progress cmdlet.
#>
Function SleepTimer-( [string]$Message, [int]$Duration )
{#############################################################################
	$intCountProcessedL = 0
	$intCountTotalL     = $Duration
	$strActivityL       = $Message
	$strMsgG            = ' '

	For(	$int1L = 1				;	`
			$int1L -le $Duration	;	`
			$int1L ++					`
		)
	{	Start-Sleep -Seconds 1
		$intCountProcessedL ++

		#Prepare progess elements
		$dblPercentageL  = 100*( $intCountProcessedL / $intCountTotalL )
		If( $dblPercentageL -gt 100 ){ $dblPercentageL = 100 }
		$intPercentageL  = [int]$dblPercentageL
		#Display progress
		Write-Progress											`
				-Activity         $strActivityL					`
				-PercentComplete  $intPercentageL				`
				-CurrentOperation "$intPercentageL% complete"	`
				-Status           $strMsgG
	}#Next
	Write-Progress -Completed
}#End Sleep Timer ############################################################


##############################################################################
Function WriteEventLogEntry-()
{#############################################################################
Param(	#Computer name if writing to remote computer.					`
		[string]$ComputerName = $Null								,	`
		#Name of EventLog to write to.									`
		[string]$LogName = 'Application'							,	`
		#Source, typically app/script name, generating entry.			`
		[string]$Source = $Null										,	`
		#Number of the EventID to be written.							`
		[int]$EventID = 0											,	`
		#Type of entry.												,	`
		[ValidateSet(	'Information', 'Warning', 'Error'			,	`
						'SuccessAudit', 'FailureAudit' ) ]				`
		[string]$EntryType = 'Information'							,	`
		#Detailed message text.											`
		[string]$Message = 'If problem persists, have a drink...'		`
		)
	If( ( -NOT $Source ) -OR ( $Source -eq '' ) )
	{	$strScriptFqL = $MyInvocation.ScriptName
		$strFileNameL = Split-Path -Path $strScriptFqL -Leaf
		$arrTempL     = $strFileNameL.Split( '.' )
		#Assume only one . in file name
		$Source = $arrTempL[0]
	}#End If - no Source

	#Create hashtable for New-EventLog use.
	$htEventDataL = @{	LogName = $LogName				;	`
						Source  = $Source				;	`
						EA      = 'SilentlyContinue'		`
						}
	#This is executed for every entry.
	#Only required before first entry.
	#Will cause an error except on first run.  Use SilentlyContinue.
	#You can only register a Source value to one LogName.
	#  If you need to write to more than one LogName, use different
	#    source names for each.
	New-EventLog @htEventDataL

	#Add the rest of the elements.
	$htEventDataL.EventID   = $EventID
	$htEventDataL.EntryType = $EntryType
	$htEventDataL.Message   = $Message

	#Only include Computer name if it is not null.
	#Including a null Computer name will cause failure.
	If( $ComputerName )
	{	$htEventDataL.ComputerName = $ComputerName }

	$htEventDataL.EA = 'Continue'
	Write-EventLog @htEventDataL
}# End Write Event Log Entry #################################################


##############################################################################
Function WriteBanner-()
{#############################################################################
[CmdletBinding()]
Param( $Nothing )
	GetCallerVariable-	-Cmdlet			$PSCmdlet							`
						-SessionState	$ExecutionContext.SessionState		`
						-VarNames		'strSeparatorG'					,	`
										'VersionInfo'					,	`
										'blnLogInitializedG'			,	`
										'strLogFileFqG'
	$strMsgL =	"`r`n`r`n"		+	`
				$strSeparatorG	+	`
				'  '			+	`
				$VersionInfo	+	`
				"`r`n    "		+	`
				( Get-Date )	+	`
				"`r`n"			+	`
				$strSeparatorG
	Write-Output $strMsgL
	If( $blnLogInitializedG )
	{	WriteLog-	-InputText $strMsgL			`
					-FilePath  $strLogFileFqG
	}#End If
}#End Write Banner ###########################################################


##############################################################################
Function WriteBoth-()
{#############################################################################
[CmdletBinding()]
Param(	[string]$InputText  = ' '	,	`
		[string]$Color      = $Null	,	`
		[switch]$Red 				,	`
		$blnLogInitializedG = $Null	,	`
		$strLogFileFqG      = $Null		`
		)
	If( -NOT $strLogFileFqG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'strLogFileFqG'
	}#End If - null
	If( -NOT $blnLogInitializedG )
	{	GetCallerVariable-	-Cmdlet       $PSCmdlet								`
							-SessionState $ExecutionContext.SessionState		`
							-VarNames     'blnLogInitializedG'
		If( -NOT $blnLogInitializedG ){ $blnLogInitializedG -eq $False }
	}#End If - null
	If( $Red )
	{	$Color = 'Red'
		Write-Host -Object $InputText -ForegroundColor $Color
	}#End If - Red
	Else
	{	If( $Color )
		{	$arrTempL = $Color.Split( '|' )
			If( $arrTempL.Length -gt 1 )
			{	#Color combo w/ background specified
				$Color    = $arrTempL[0]
				$strBackL = $arrTempL[1]
				Write-Host	-Object          $InputText	`
							-ForegroundColor $Color		`
							-BackgroundColor $strBackL
			}#End If - two colors
			Else{ Write-Output $InputText -ForegroundColor $Color }
		}#End - Color
		Else{ Write-Host -Object $InputText }
	}#End Else - not Red

	If( $blnLogInitializedG )
	{	WriteLog-	-InputText $InputText		`
					-FilePath  $strLogFileFqG
	}#End If - initialized
}#End Write Both #############################################################


##############################################################################
Function WriteLog-()
{#############################################################################
Param(	$InputText = ''											,	`
		[ Parameter( Mandatory = $True ) ][string]$FilePath		,	`
		[switch]$NewFile											`
		)
	If( $NewFile )
	{	$InputText | Out-File	-FilePath $FilePath	`
								-Encoding "ASCII"
	}#End If
	Else
	{	$InputText | Out-File	-FilePath $FilePath	`
								-Encoding "ASCII"	`
								-Append
	}#End Else - appending
}#End Write Log ##############################################################


##############################################################################
Function WriteVerbose-()
{#############################################################################
[CmdletBinding()]
Param(	[string]$InputText				,	`
		[string]$Color      = 'Blue'	,	`
		[switch]$NoHost					,	`
		[switch]$NoLog					,	`
		$blnLogInitializedG = $Null		,	`
		$strLogFileFqG      = $Null			`
		)
	If( $VerbosePreference -eq 'Continue' )
	{	$InputText     = "<V> $InputText"
		$arrTempL      = $Color.Split( '|' )
		$strForeColorL = $arrTempL[0]

		If( $arrTempL.Length -gt 1 )
		{	#Background color specified
			$strBackColorL = $arrTempL[1]
		}#End If
		Else{	$oBackColorL = (Get-Host).ui.rawui.BackgroundColor
				$strBackColorL = [string]$oBackColorL
			}#End Else

		If( -NOT( $NoHost ) )
		{	Write-Host $InputText	-ForegroundColor $strForeColorL	`
									-BackgroundColor $strBackColorL
		}#End If
		If( $NoLog -eq $False )
		{	#Writing to log...
			If(	( -NOT $blnLogInitializedG ) )
			{	GetCallerVariable-												`
							-Cmdlet			$PSCmdlet							`
							-SessionState	$ExecutionContext.SessionState		`
							-VarNames		'blnLogInitializedG'
			}#End If - blnLogInitializedG
			If(	( -NOT $strLogFileFqG ) )
			{	GetCallerVariable-												`
							-Cmdlet			$PSCmdlet							`
							-SessionState	$ExecutionContext.SessionState		`
							-VarNames		'strLogFileFqG'
			}#End If - strLogFileFqG

			If( $blnLogInitializedG )
			{	WriteLog-	-InputText $InputText			`
							-FilePath  $strLogFileFqG
			}#End If - initialized
		}#End If - not NoLog
	}#End If - verbose continue
}#End Write Verbose ##########################################################
