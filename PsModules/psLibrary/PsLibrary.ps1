##############################################################################
Function GetLibraryVersion_(){ Return 'WorflowLibrary Version - 1.0.1' }
# 1.0.1 - Added switches to WriteVerbose_
##############################################################################


##############################################################################
Function BinaryCodedValues_()
{#############################################################################
Param(	[int]$InputValue	,	`
		[array]$ValueNames		`
		)
#Returns an array of strings coresponding to their binary values in a "bitmask"
#  integer value.
#If there is not a string for each numeric value, pad the input array
#  with a filler string for any gaps.
#Don't need high values if not used, but must have contiguous values
#  up to last one used.
	$arrThresholdsL   = @( 1,2,4,8,16,32,64,128,256,512,1024 )
	$intCurrentValueL = $InputValue
	$colResultsL      = @()
	$strMsgL          =	'  '				+	`
						$ValueNames.Count	+	`
						' possible values.'

	For(	$intThresholdIndexL = $arrThresholdsL.GetUpperBound(0)	;	`
			$intThresholdIndexL -ge 0								;	`
			$intThresholdIndexL --										`
		)
	{	$intThresholdValueL = [int]( $arrThresholdsL[ $intThresholdIndexL ] )
		$strMsgL =	'  Threshold index - '					+	`
					$intThresholdIndexL						+	`
					"`r`n"									+	`
					'  Evaluating, '						+	`
					$intCurrentValueL						+	`
					', against threshold value, '			+	`
					$intThresholdValueL
		WriteVerbose_ $strMsgL

		If( $intCurrentValueL -ge $arrThresholdsL[ $intThresholdIndexL ] )
		{	$strValueToAddL = $ValueNames[ $intThresholdIndexL ]
			$strMsgL =	'  Value is enabled.'	+	`
						"`r`n"					+	`
						'  Adding, '			+	`
						$strValueToAddL
			WriteVerbose_ $strMsgL
			$colResultsL      += $strValueToAddL
			$intCurrentValueL -= $intThresholdValueL
		}#End If
	}#Next - intThresholdIndexL
	Return $colResultsL
}#End Binary Coded Vaulues ###################################################


##############################################################################
Function DateTimeString_()
{#############################################################################
[cmdletbinding()]
Param(	[dateTime]$DateTime = ( Get-Date )							,	`
		[ValidateSet( '_', '-', ' ' )] [string]$Separator  = '_'	,	`
		[switch]$Date												,	`
		[switch]$Time												,	`
		[switch]$Milliseconds											`
		)
	If( ( -NOT( $Date ) ) -AND ( -NOT( $Time ) ) ){ Return '' }
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
		{	$strTempL = [string]($DateTime).Millisecond }
		If( $strTempL.Length -eq 1 ){ $strTempL = "00$strTempL" }
		If( $strTempL.Length -eq 2 ){ $strTempL = "0$strTempL" }
		$strTimeStringL += $strTempL
	}#End If - Time

	If( ( $Date -AND $Time ) )
	{ Return $strDateStringL + $Separator + $strTimeStringL }
	Else
	{ Return $strDateStringL + $strTimeStringL }
}#End Date Time String #######################################################


##############################################################################
Function Decrypt_( [string]$strTextP, [string]$strKeyP)
{#############################################################################
	$arrTextL  = @()
	$arrInputL = $strTextP.ToCharArray()
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
	{	$strKeyCharL   = $strKeyP.Substring( $intKeyCounterL, 1 )
		$intClearCharL = $intCharL - [Int][Char]$strKeyCharL
		$strOutL      += [Char]$intClearCharL
		$intKeyCounterL ++
		If( $intKeyCounterL -eq $strKeyP.Length ){ $intKeyCounterL = 0 }
	}#Next
	Return $strOutL
}#End Decrypt ################################################################


<#############################################################################
.Synopsis
Simplifies generation of elapsed time string.

.Description
Default call, no arguments, starts interval with current time.
Call with -Begin value starts interval using passed value.
Call with -Stop switch terminates interval using current time.
Call with -End value terminates interval using passed value.
Can call with both -Begin and ( -End value or -Stop switch ).
#>
Function ElapsedTime_()
{#############################################################################
Param(	[dateTime]$Begin	,	`
		[dateTime]$End		,	`
		[switch]$Stop			`
		)
	#No arguments
	If(	$Begin -eq $Null	`
		-AND				`
		$End -eq $Null		`
		-AND				`
		$Stop -eq $False	`
		)
	{	#Called with no arguments
		$Begin = Get-Date
	}#End If - no args

	If( $Begin )
	{	$Script:EtBeginG = $Begin
		$strElapsedL     = 'Begin'
		$strMsgL         = "ET interval begins, $Begin."
		WriteVerbose_ $strMsgL
		If( $Stop )
		{	If( $End -eq $Null ){ $End = Get-Date }
			$strElapsedL = ETString_	-Start $Begin	`
										-Stop  $End
			$strMsgL = "ET interval ends, $End, $strElapsedL."
			WriteVerbose_ $strMsgL
		}#End If - Stop

		If( $End )
		{	#No error if both -Begin, -Stop and -End are used.
			#-End time used with -Stop
			$strElapsedL = ETString_	-Start $Begin	`
										-Stop  $End
			$strMsgL = "ET interval ends, $End, $strElapsedL."
			WriteVerbose_ $strMsgL
		}#End If - Stop
	}#End If - Begin has value
	Else
	{	If( $Stop -OR $End )
		{	#Called with either -Stop switch or $End value.
			#-Stop switch will use current time.
			If( $End -eq $Null ){ $End = Get-Date }

			If( $Script:EtBeginG )
			{	#Beginning time exists
				$strElapsedL = ETString_	-Start $Script:EtBeginG	`
											-Stop  $End
				#Null this value as prep for next Begin operation
				$Script:EtBeginG = $Null
				$strMsgL = "ET interval ends, $End, $strElapsedL."
				WriteVerbose_ $strMsgL
			}#End If - EtBeginG exists
			Else
			{	$strMsgL =	'**Error - Called ElapsedTime_() -Stop w/o Start'
				WriteBoth_ $strMsgL -Red
				ExitProcessing_
			}#End Else - Err, no begin value
		}#End If - Stop or End
	}#End If - Not Begin

	Return $strElapsedL
}# End Elapsed Time ##########################################################


##############################################################################
Function Encrypt_( [string]$strTextP, [string]$strKeyP)
{#############################################################################
	$arrTextL       = $strTextP.ToCharArray()
	$intKeyCounterL = 0
	ForEach( $strCharL IN $arrTextL )
	{	$strKeyCharL   = $strKeyP.Substring( $intKeyCounterL, 1 )
		$intCryptCharL = [Int][Char]$strCharL + [Int][Char]$strKeyCharL
		$strCharOutL   = [string]$intCryptCharL
		While( $strCharOutL.Length -lt 3 ){ $strCharOutL = '0' + $strCharOutL }
		$strOutL      += $strCharOutL
		$intKeyCounterL ++
		If( $intKeyCounterL -eq $strKeyP.Length ){ $intKeyCounterL = 0 }
	}#Next
	Return $strOutL
}#End Encrypt ################################################################


##############################################################################
Function EndChar_()
{#############################################################################
Param(	[string]$String			,	`
		[string]$EndChar = '\'		`
		)
	If( $String.EndsWith( $EndChar ) ){ Return $String }
	Else{ Return $String + $EndChar }
}#End End Char ##############################################################


##############################################################################
Function EndColon_( $Input )
{#############################################################################
	Return EndChar_ -String $Input -EndChar ':'
}#End End Colon ##############################################################


##############################################################################
Function EndSemicolon_( $Input )
{#############################################################################
	Return EndChar_ -String $Input -EndChar ':'
}#End End Semicolon ##########################################################


#############################################################################
Function ErrorHandler_()
{############################################################################
Param(	[string]$Description = $Script:ErrorDescriptionG	,	`
		[int]$Index = 0											,	`
		[switch]$Abort												`
		)
#$Index is reference to specific error, with zero being current.
#Higher values refer to preceding errors.

	$strExMessageL    = Paragraph_ $Error[$Index].Exception.Message
	$strCategoryInfoL = Paragraph_ $Error[$Index].CategoryInfo
	$strFqErrorIdL    = Paragraph_ $Error[$Index].FullyQualifiedErrorID
	$strScriptNameL   = Paragraph_ $Error[$Index].InvocationInfo.ScriptName
	$strInvocInfoL    = Paragraph_ $Error[$Index].InvocationInfo.Line
	$strLineInfoL     =	[string]$Error[0].InvocationInfo.ScriptLineNumber	+	`
						" / "												+	`
						$Error[0].InvocationInfo.OffsetInLine

	$strMsgL =	$strSeparatorG				+	`
				'**An error has occurred.'	+	`
				"`r`n"						+	`
				'    Exception: '			+	`
				"`r`n"						+	`
				$strExMessageL				+	`
				"`r`n"						+	`
				'    Category:'				+	`
				"`r`n"						+	`
				$strCategoryInfoL			+	`
				"`r`n"						+	`
				'    Full ID:'				+	`
				"`r`n"						+	`
				$strFqErrorIdL				+	`
				"`r`n"						+	`
				'    Script:'				+	`
				"`r`n"						+	`
				$strScriptNameL				+	`
				"`r`n"						+	`
				"`r`n"						+	`
				'    Line/Char:'			+	`
				"`t"						+	`
				$strLineInfoL				+	`
				"`r`n"						+	`
				'    Failed command:'		+	`
				"`r`n"						+	`
				$strInvocInfoL				+	`
				"`r`n"						+	`
				'    Note:'					+	`
				"`r`n`t"						+	`
				$Description				+	`
				"`r`n"						+	`
				$strSeparatorG
	WriteBoth_ $strMsgL
	If( $Abort ){ ExitProcessing_ }
}#End Error Handler ##########################################################


##############################################################################
Function EtString_()
{#############################################################################
Param(	[DateTime]$Start				,	`
		[DateTime]$Stop = (Get-Date)		`
		)
	If( $Start -eq $Null )
	{	WriteBoth(	"A null value was passed to the EtString "	+	`
					"function as the starting time..." )
		ExitProcessing_
	}#End If
	$strElapsedTimeL = ""
	$strTempL        = ""
	$ElapasedL       = New-TimeSpan $Start $Stop

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
		$strMSecsL = PadString_ $strMSecsL 3 -PadChar "0" -Left

		If( $strElapsedTimeL.Length -gt 0 )
		{	$strElapsedTimeL += ", " }
		$strTempL = [string]$ElapasedL.Seconds	+	`
					"."							+	`
					$strMSecsL
		$strElapsedTimeL += "$strTempL seconds"
	}#End If - Seconds
	Return $strElapsedTimeL
}#End ET String ##############################################################


##############################################################################
Function ExitProcessing_( [int]$ReturnCode = 0 )
{#############################################################################
	$strTimeL = ( Get-Date ).ToString()
	$strMsgL =	"`r`n"						+	`
				$Script:strSeparatorG		+	`
				"Processing complete...`t"	+	`
				$strTimeL					+	`
				"`r`n"
	If( $Script:strLogFileFqG -ne '' )
	{	$strMsgL +=	'  Detailed information may be found in log file:'	+	`
					"`r`n`t"											+	`
					$Script:strLogFileFqG								+	`

					"`r`n`r`n`r`n"
	}#End If
	#$strMsgL += "`r`nDone."
	WriteBoth_ $strMsgL

	Exit $ReturnCode
}#End Exit Processing ########################################################


##############################################################################
Function ExtractServerName_( [string]$PathString )
{#############################################################################
	If( $PathString -Match ':' )
	{	#Assume MS-DOS path on local machine
		$strServerL = $Env:ComputerName
	}#End If
	Else
	{	While( $PathString.StartsWith( '\' ) )
		{	$PathString = $PathString.Substring( 1 ) }
		$arrTemp    = $PathString.Split( '\' )
		$strServerL = $arrTemp[0]
	}#End Else
	Return $strServerL
}#End Extract Server Name ####################################################


##############################################################################
Function GetCredNames_()
{#############################################################################
#Looks for file holding Domain\UserName for each credential.
#If found, subsitutes for global vars.  Then credential processing continues.
	If( Test-Path $Script:RunAsUserNameFile )
	{	#File exists with user names for credentials.
		$arrCredNamesL = Get-Content $Script:RunAsUserNameFile
		ForEach( $strLineL IN $arrCredNamesL )
		{	$strLineL = $strLineL.Trim()
			If( $strLineL.Length -gt 0 -AND ! $strLineL.StartsWith( '#' ) )
			{	$arrLineL = $strLineL.Split( ',' )
				Switch( $arrLineL[0] )
				{	'AdLocal'  { $Script:RunAsUserAD       = $arrLineL[1] }
					'AdRemote' { $Script:RunAsUserAdTarget = $arrLineL[1] }
					'ExLocal'  { $Script:RunAsUserExLocal  = $arrLineL[1] }
					'ExRemote' { $Script:RunAsUserExRemote = $arrLineL[1] }
					'MDB'      { $Script:RunAsUserMDB      = $arrLineL[1] }
				}#End switch
			}#End If
		}#Next - line
	}#End If - file exists
}#End Get Credential Names ###################################################


##############################################################################
Function GetFolderSize_()
{#############################################################################
Param(	[string]$UNC	,	`
		[switch]$Test		`
		)

	If( $Test )
	{	If( -NOT( Test-Path $UNC ) )
		{	$strMsgL =	'**Error - Invalid UNC Path'	+	`
						"`r`n  "						+	`
						$UNC							+	`
						', not found.'
			WriteBoth_ $strMsgL -Red
			ExitProcessing_
		}#End If - Test path
	}#End If - Test

	$colFilesL = Get-ChildItem $UNC -Recurse

	$dblTotalSizeL = 0
	ForEach( $oFileL IN $colFilesL )
	{	$strAttributesL = $oFileL.Attributes.ToString().ToLower()

		If( $strAttributesL.Contains( 'directory' ) )
		{	$blnIsDirectoryL = $True
			$strMsgL = '  Folder - ' + $oFileL.Name
			#WriteVerbose_ $strMsgL
		}#End If
		Else
		{	$dblTotalSizeL += $oFileL.Length
			$strMsgL =	'  File - '			+	`
						$oFileL.FullName	+	`
						' - '				+	`
						$oFileL.Length
			#WriteVerbose_ $strMsgL
		}#End If
	}#Next - File
	Return $dblTotalSizeL
}#End - Get Folder Size ######################################################


<#############################################################################
.Synopsis
Takes a string argument and examines it to see if it is single item,
  a comma-separated list, or a file name.  Then it creates an array from the
  item, comma-separated list or the contents of the file.

.Description
This calls the ReadInputTextFile_() function, but does not offer the option
  of returning a hashtable, which can be done calling the function directly.
#>
Function GetList_()
{#############################################################################
Param(	[string]$List							,	`
		[string]$Delimiter = $Script:Delimiter		`
		)
	$arrListL = @()

	If( $List.Contains( '.' ) )
	{	$arrTempL = $List.Split( '.' )
		If( $arrTempL.GetUpperBound(0) -eq 0 )
		{	#Only one period.
			#Assume it is a file name.
			If( -NOT( $List.Contains( '\' ) ) )
			{	#File name only, not full path.
				$List = $Script:ScriptPath + $List
			}#End If - path
		}#End If - file name
		$arrListL = ReadInputTextFile_ $List
	}#End If - .
	Else
	{	#List is a delimited list of items
		$arrListL = $List.Split( $Delimiter )
	}#End Else

	If( $Script:Verbose )
	{	$strMsgL = '  The following items will be processed:'
		ForEach( $strItemL IN $arrListL )
		{	$strMsgL += "`r`n    $strItemL" }
		WriteBoth_ $strMsgL
	}#End If - Verbose
	Return $arrListL
}#End Get List ###############################################################


##############################################################################
Function GetQadObjectType_( [string]$Input )
{#############################################################################
	$oObjectL = Get-QadObject $Input
	If( $oObjectL -eq $null ){ Return 'Empty' }
	Return $oObjectL.Type
}#End Get Object Type ########################################################


##############################################################################
Function GetServerList_( [string]$List )
{#############################################################################
	$arrListL = @()

	If( $List )
	{	If( $List -eq 'Database' )
		{	#Get list from database
			$strQueryL    = "Select * From sys_Servers"
			$oServerDataL = SqlInitializeDataset_ $strQueryL

			ForEach( $oServerL in $oServerDataL.Tables[0].Rows )
			{	$strServerNameL  = [string]$oServerL.ServerName
				$arrListL += $strServerNameL
			}#Next - oServerL
			###$oServerDataL.Close()
			###Return $arrListL
		}#End If - Database
		Else
		{	#Not calling database, process as list/file
			$arrListL = GetList_ $List
			Return $arrListL
		}#End Else - not database
	}#End If - List not null
	Else
	{	$strMsgL = 'No server name/list specified.  Using local computer.'
		WriteBoth_ $strMsgL
		$arrListL = @( $Env:ComputerName )
	}#End Else
	If( $Script:Verbose )
	{	$strMsgL = '  The following will be processed:'
		ForEach( $strServerL IN $arrListL )
		{	$strMsgL += "`r`n    $strServerL" }
		WriteBoth_ $strMsgL
	}#End If - Verbose
	Return $arrListL
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
Function GetSystemInfoSQL_( [int]$ConfigNumber = 1 )
{#############################################################################
#Retrieves configuration items from database.
#Uses Script scope datareader
	$strQueryL =	'Select * From sys_SystemInformation '	+	`
					"WHERE ConfigID = $ConfigNumber"
	SqlInitializeReader_ $strQueryL
	While( $Script:oSqlReaderG.Read() )
	{	#Create SMTP Server Object
		$Script:oSmtpServerG = New-Object System.Object
		$Script:oSmtpServerG | Add-Member										`
								-Type  NoteProperty								`
								-Name  ServerName								`
								-Value $Script:oSqlReaderG['SmtpServer'].ToString()
		$Script:oSmtpServerG | Add-Member										`
								-Type  NoteProperty								`
								-Name  Port										`
								-Value $Script:oSqlReaderG['SmtpPort']
		$Script:oSmtpServerG | Add-Member										`
								-Type  NoteProperty								`
								-Name  UseAuth									`
								-Value $Script:oSqlReaderG['SmtpUseAuth']
		$Script:oSmtpServerG | Add-Member										`
								-Type  NoteProperty								`
								-Name  AuthName									`
								-Value $Script:oSqlReaderG['SmtpAuthName'].ToString()
		$Script:oSmtpServerG | Add-Member										`
								-Type  NoteProperty								`
								-Name  AuthPwd									`
								-Value $Script:oSqlReaderG['SmtpAuthPwd'].ToString()
	}#WEnd
	$Script:oSqlReaderG.Close()
}#End Get System Info SQL ####################################################


<#############################################################################
.Synopsis
Returns an array of two elements.
First element is a string indicating the data type of second element.

.Description
$TimeInput argument is a string.
If it appears to be a 1-2 digit integer, it is returned as an integer.
The two digit value probably used as the hour for a recurring/continuous run.
If it appears to be a dateTime value, it is returned as such.
The dateTime format probably used for a one-time, not recurring run.
#>
Function GetTime_( [string]$TimeInput )
{#############################################################################
	If( ( $Start.Length -ge 1 ) -AND ( $Start.Length -le 2 ) )
	{	#Probably an hour
		$strEaHoldL            = $ErrorActionPreference
		$ErrorActionPreference = 'SilentlyContinue'
		$intHourNowL              = [int]$TimeInput

		If( !( $? ) )
		{	$strMsgL =	'**Error - Invalid dateTime string format'	+	`
						"`r`n  String value, "						+	`
						$TimeInput									+	`
						', could not be converted to an hour value.'
			WriteBoth_ $strMsgL -Red
			ExitProcessing_
		}#End If - Error

		Return @( 'Integer', $intHourNowL )
	}#End If - 2 digits

	#DateTime string format: 8/20/2014 11:10:07 AM
	#Most partials of the full format will be converted successfully.
	$strEaHoldL            = $ErrorActionPreference
	$ErrorActionPreference = 'SilentlyContinue'
	$dteReturnL            = [dateTime]$TimeInput

	If( !( $? ) )
	{	$strMsgL =	'**Error - Invalid dateTime string format'				+	`
					"`r`n  String value, "									+	`
					$TimeInput												+	`
					"`r`n    could not be converted to a dateTime value."
		WriteBoth_ $strMsgL -Red
		ExitProcessing_
	}#End If - Error

	$ErrorActionPreference = $strEaHoldL
	Return @( 'DateTime', $dteReturnL )
}# End Get Time ##############################################################


##############################################################################
Function InitializeLog_()
{#############################################################################
#Initializes logging based on script's command line parms.
# -LogAppend controls overwrite of existing file of same name.
# -LogDate, -LogTime control date/timestamp in file name.
##############################################################################
	$blnFileExistsL = $False
	$strDateStringL = ''

	If( $Script:LogFolder -eq '' )
	{	WriteVerbose_ 'No log folder specified, using script folder.'
		$Script:LogFolder = $Script:strScriptPathG
	}#End If - ''
	$Script:LogFolder = EndBackslash( $Script:LogFolder )

	If( -NOT ( Test-Path $Script:LogFolder ) )
	{	$strMsgL =	'**Error - Log folder is not valid.'	+	`
					"`r`n`t$Script:LogFolder"
		Write-Host $strMsgL
		ExitProcessing_ 77
	}#End If

	#Construct date/time portion of file name
	If($Script:LogDate -AND $Script:LogTime )
	{	$strDateStringL = DateTimeString_ -Date -Time }
	Else
	{	If( $Script:LogDate ){ $strDateStringL = DateTimeString_ -Date }
		If( $Script:LogTime ){ $strDateStringL = DateTimeString_ -Time }
	}#End Else - date + time

	$Script:strLogFileFqG =	$Script:LogFolder	+	`
							$Script:LogBase		+	`
							$strDateStringL		+	`
							".txt"
	#Delete existing log file, if needed
	If( -NOT $Script:LogAppend )
	{	If( Test-Path( $Script:strLogFileFqG ) )
		{	Remove-Item $Script:strLogFileFqG }
	}#End If - Test Path
	$Script:blnLogInitializedG = $True

	If( $Script:arrLogHoldG.GetUpperBound(0) -gt 0 )
	{	For(	$int1L = 1											;	`
				$int1L -le $Script:arrLogHoldG.GetUpperBound( 0 )	;	`
				$int1L ++ )
		{	WriteLog_ $Script:arrLogHoldG[ $int1L ] }
	}#End If - Content in array
}#End Initialize Log #########################################################


##############################################################################
Function IsAdmin_()
{#############################################################################
	$UserL = [Security.Principal.WindowsIdentity]::GetCurrent()
	(New-Object Security.Principal.WindowsPrincipal $UserL).IsInRole(	`
			[Security.Principal.WindowsBuiltinRole]::Administrator)
}#End Is Admin ###############################################################


##############################################################################
Function IsDbNull_( $InputP )
{#############################################################################
	Return  [System.DBNull]::Value.Equals($InputP)
}#End Is DbNull ##############################################################


##############################################################################
Function ListObjectProperties_( $Object )
{#############################################################################
	$strMsgL = "`r`n  Type name is " + $Object.GetType().Fullname
	Write-Host $strMsgL

	$colPropertiesL = $Object.PsObject.Properties

	If( !( $colPropertiesL ) )
	{	Write-Host "  Object has no properties"
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
		Write-Host $strMsgL
	}#Next
}#End List Object Properties #################################################


##############################################################################
Function LoadEx2007Snapin_( [switch]$NoAbort )
{#############################################################################
	$strSnapinL = 'Microsoft.Exchange.Management.PowerShell.Admin'
	Add-PsSnapin $strSnapinL -EA 'SilentlyContinue'
	$oSnapinL = Get-PSSnapin $strSnapinL
	If( $oSnapinL ){ Return $True }
	Else
	{	$strMsgL =	'**Error - Unable to load Exchange 2007 snapin.'
		WriteBoth_ $strMsgL -Red
		If( $NoAbort ){ Return $False }
		Else { ExitProcessing_ }
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
LoadModule_ -ModuleName 'poaShell'  -ModulePath 'C:\Temp'
	Checks to see if poaShell is loaded/available and attempts to copy
	  and load it if not.  Aborts on failure.
#>
Function LoadModule_()
{#############################################################################
Param(	#Name of module to loaded									`
		[Parameter ( Mandatory = $True ) ]							`
		[string]$ModuleName										,	`
		#Path to source holding module to install and load			`
		[string]$ModulePath = $Null								,	`
		#Switch to only check load status no attempt to load		`
		[switch]$CheckOnly										,	`
		#Switch to allow continue after failure to load				`
		[switch]$NoAbort											`
		)
	$strMsgL = "  Examining status of module: $ModuleName..."
	WriteVerbose_ $strMsgL

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
				WriteVerbose_ '  Module is loaded'
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
		WriteVerbose_ '  Module is not loaded.  Checking available...'
		$colLoadedModulesL = Get-Module -ListAvailable
		If( $colLoadedModulesL )
		{	ForEach( $oLoadedModuleL IN $colLoadedModulesL )
			{	$strMsgL =	'  Loaded module: '	+	`
							$oLoadedModuleL.Name
				WriteVerbose_ $strMsgL

				If( $oLoadedModuleL.Name -eq $ModuleName )
				{	#Flag as available, load later...
					$blnIsAvailableL = $True
					Break
				}#End If - found
			}#Next - oLoadedModuleL

			If( $blnIsAvailableL -eq $False )
			{	#Module was not loaded AND is not available.
				If( $Script:ModulePath = $Null )
				{	#No source path was specified to install module.
					$strMsgL =	'**Warning - specified module not '	+	`
								'loaded OR available.'
					WriteBoth_ $strMsgL -Color 'Yellow|Black'
					Return $False
				}#End If - no ModulePath
			}#End If - not available
		}#End If - some module available
	}#End If - Available

	If( -NOT( $blnIsAvailableL ) )
	{	WriteVerbose_ "`r`n  Module is not available"
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
		WriteVerbose_ $strMsgL

		#Modules folder may not exist
		#Not created until first module installed/loaded
		$strModulesPathL   = $strProfileFolderL + 'Modules\'

		#Check for Modules Folder
		If( Test-Path $strModulesPathL )
		{	WriteVerbose_ "`r`n  Local Modules folder exists" }
		Else
		{	WriteVerbose_ "`r`n  Modules folder does not exist.  Creating it..."
			$oModulesPathL = New-Item $strModulesPathL -Type Directory
		}#End If - no modules folder

		#Target folder for specified module.
		$strTargetFolderL = $strModulesPathL + $ModuleName

		#Check for target module folder
		If( Test-Path $strTargetFolderL )
		{	$strMsgL =	"`r`n  Local Module folder exists: "	+	`
						"`r`n`t$strTargetFolderL"
			WriteVerbose_ $strMsgL
			#Folder exists.  Assume module in folder.
			$blnIsAvailableL = $True
		}#End If - Test-Path
		Else
		{	#Folder for module is not found in profile.
			#See if a module source location is specified.
			If( $ModulePath )
			{	#A source folder for the module is specified.
				#Module will be copied from there to local profile.
				$strSourceFolderL  = EndBackSlash $ModulePath
				$strSourceFolderL += $ModuleName
				$strMsgL =	"`r`n  Local module folder does not exist: `r`n`t"		+	`
							$strTargetFolderL										+	`
							"`r`n  Will copy from source,`r`n`t$strSourceFolderL"
				WriteVerbose_ $strMsgL
				If( Test-Path $strSourceFolderL )
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
						WriteBoth_ $strMsgL -Red
						ErrorHandler_ "Failed to copy module, $ModuleName"
						If( $NoAbort ){ Return $False }
						Else{ ExitProcessing_ 78 }
					}#End If - error
				}#End If - source path exists
				Else
				{	$strMsgL =	'**Error - Source path for module '	+	`
								'installation does not exist'		+	`
								"`r`n`t$strSourceFolderL"
					WriteBoth_ $strMsgL -Red
					ExitProcessing_ 79
				}#End Else - source not found
			}#End Else - Module now copied to destination
			Else
			{	$strMsgL =	'**Error - Specified PS module not '	+	`
							'installed or available locally'		+	`
							"`r`n`tModule: $ModuleName"				+	`
							"`r`n  No alternate source for "		+	`
							'installation was specified.'
				WriteBoth_ $strMsgL -Red
				ExitProcessing_ 79
			}#End Else - Module source not specified.
		}#End If - module source path specified
	}#End If - not avaialable
	Else
	{	#Module available, try to load it
		$strEaHoldL            = $ErrorActionPreference
		$ErrorActionPreference = 'SilentlyContinue'
		WriteVerbose_ '  Loading module...'
		Import-Module $ModuleName

		If( $? )
		{	Return $True }
		Else
		{	ErrorHandler_ "Failed to load module, $ModuleName"
			If( $NoAbort ){ Return $False }
			Else{ ExitProcessing_ 77 }
		}#End If - error
		$ErrorActionPreference = $strEaHoldL
	}#End Else - available
	Return $False	#Shouldn't actually get to here...
}# End Load Module ###########################################################


##############################################################################
Function MakeUNC_()
{#############################################################################
Param(	$Computer = $Script:Computer		,	`
		$Path     = $Script:Path		`
		)

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
		WriteBoth_ $strMsgL -Red
		ExitProcessing_
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
	{	$strMsg =	'  Item '			+	`
					$int1L				+	`
					': '				+	`
					$arrTempL[ $int1L ]
		$strUncL += '\' + $arrTempL[ $int1L ]
	}#Next
	Return $strUncL
}#End - Make UNC #############################################################


##############################################################################
Function MdbConnect_()
{#############################################################################
Param(	$DB								,	`
		$Path	=  $Script:MdbFolder	,	`
		$User	= "Admin"				,	`
		$Pwd	= ""						`
		)
	$strTargetDbL = ( EndBackslash $Script:MdbFolder ) + $DB
	If( _NOT Test-Path $strTargetDbL )
	{	$strMsgL = '**Error - specified database file not found.'
		WriteBoth_ $strMsgL -Red
		ExitProcessing_
	}#End If

	$strConnect32L	= "Driver={Microsoft Access Driver (*.mdb)};"
	$strConnect64L	= "Driver={Microsoft Access Driver (*.mdb, *.accdb)};"

	$Path = EndBackslash $Path
	$strConnect2L =	"Dbq="			+	`
					$DB				+	`
					";DefaultDir="	+	`
					$Path			+	`
					";Uid="			+	`
					$User			+	`
					";Pwd="			+	`
					$Pwd			+	`
					";"
	WriteVerbose_ $strConnect2L

	#Store Error action preference
	$strErrorActionHoldL = $ErrorActionPreference
	$ErrorActionPreference = "SilentlyContinue"

	###Trap{ ErrorHandler_ }
	$Script:ConnectionG = New-Object -ComObject ADODB.Connection

	#Attempt 32 bit connect first
	$strCStringL = $strConnect32L + $strConnect2L
	$Script:ConnectionG.Open( $strCStringL )

	If( -NOT $? )
	{	#Connection attempt failed, try 64 bit
		$strCStringL = $strConnect64L + $strConnect2L
		$Script:ConnectionG.Open( $strCStringL )

		If( -NOT $? )
		{	#Second attempt failed
			$strTempL =	"  Both 64 bit and 32 bit connection attempts "	+	`
						"to connect to database, "						+	`
						$DB												+	`
						", failed.  This is the error for the 64 bit "	+	`
						"attempt."
			ErrorHandler -Description $strTempL -Index 0

			$strTempL =	"  Both 64 bit and 32 bit connection attempts "	+	`
						"to connect to database, "						+	`
						$DB												+	`
						", failed.  This is the error for the 32 bit "	+	`
						"attempt."
			ErrorHandler -Description $strTempL -Index 1
			ExitProcessing_
		}#End If - second ?
	}#End If - first ?
	#Restore error action
	$ErrorActionPreference = $strErrorActionHoldL
}#End - MDB Connect ##########################################################


##############################################################################
Function MdbSelectRecords_( [string]$strQueryP )
{#############################################################################
	$RecordSetL = New-Object -ComObject ADODB.Recordset
	WriteVerbose_ "  $strQueryP"
	Trap{ ErrorHandler_ }
	$RecordSetL.Open(	$strQueryP,				`
						$Script:ConnectionG,	`
						$Script:OpenStaticG,	`
						$Script:LockOptimisticG	`
					)
	Return $RecordsetL
}#End MDB Select Records #####################################################


##############################################################################
Function MdbUpdateRecords_()
{#############################################################################
Param(	[string]$TableName		,	`
		[string]$WhereClause	,	`
		$DataTable					`
		)
	$WhereClause = EndSemiColon $WhereClause
	$WhereClause = ' ' + $WhereClause
	$arrKeysL    = $DataTable.Keys
	$strQueryL   =	'SELECT * '	+	`
					'FROM '		+	`
					$TableName	+	`
					' '			+	`
					$WhereClause
	WriteVerbose_ $strQueryL
	$RecordSetL = New-Object -ComObject ADODB.Recordset
	$RecordSetL.Open(	$strQueryL,				`
						$Script:ConnectionG,	`
						$Script:OpenStaticG,	`
						$Script:LockOptimisticG	`
					)
	While ( $RecordSetL.EOF -ne $True  )
	{	ForEach( $strKeyL IN $arrKeysL )
		{	$RecordSetL.Fields.Item( $strKeyL ).value = $DataTable.$strKeyL }
		$RecordSetL.Update()
		$RecordSetL.MoveNext()
	}#WEnd - EoF
	$RecordSetL.Close()
}#End MDB Update Records #####################################################


##############################################################################
Function PadString_()
{#############################################################################
Param(	[string]$StringToPad	,	`
		[int]$Length			,	`
		[string]$PadChar = " "	,	`
		[switch]$Left			,	`
		[switch]$Right			,	`
		[switch]$Both			,	`
		[switch]$HTML
		)
#Accepts an input string and length (mandatory)
#Pads string with specified character to specified length.
#Padded on Left/Right/Both, favoring Left and defaulting to Left.
#This version truncates input if longer than $Length
	$intDiffL = $Length - $StringToPad.Length
	If( $intDiffL -le 0 )
	{	Return $StringToPad.Substring( 0, $Length ) }

	If( !$Left -AND !$Right -AND !$Both ){ $Left = $True }
	If( $Left )
	{	For(	$int1L = 1				;	`
				$int1L -le $intDiffL	;	`
				$int1L ++					`
			)
		{	$StringToPad = $PadChar + $StringToPad }
	}#End If - Left

	If( $Right )
	{
		For(	$int1L = 1				;	`
				$int1L -le $intDiffL	;	`
				$int1L ++					`
			)
		{	$StringToPad += $PadChar }
	}#End If - Right

	If( $Both )
	{	$intLeftL = [int]($intDiffL / 2 )
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
}#End Function - Pad String ##################################################


############################################################################\
Function Paragraph_()
{############################################################################
#Accepts a string of text and optional parameters to reformat text to
#  specified paragraph/line dimensions.
#LeftMargin is whitespace added to each line as spaces.
#IndentFirst is an optional indent for the first line, added to the value
#  of LeftMargin.
#TrimFirst removes all left-padding from first line, giving text of same
#  width as rest of lines, without left-padding.
############################################################################/
Param(	[string]$Text				,	`
		[int]$RightMargin    = 76	,	`
		[int]$LeftMargin     = 8	,	`
		[int]$IndentFirst    = 0	,	`
		[switch]$TrimFirst   		,	`
		[switch]$ReturnArray
		)
	$arrInputL        = $Text.Split( " " )
	$intLastIncludedL = 0
	$intTrimAtL       = 0
	$strExtraL        = ""
	$strLMarginPadL     = " " * $LeftMargin
	$strTextPieceL    = ""

	#Not sure why...
	If( $RightMargin - $LeftMargin -le 13 )
	{	Write-Host "  Minimum text width, LineWidth - Margin is 14 characters."
		ExitProcessing_
		Break
	}#End If width 13

	If( $TrimFirst ){ $intTextWidthL = $RightMargin }
	Else            { $intTextWidthL = $RightMargin - $LeftMargin }

	#Add indent to first word/line
	$arrInputL[0] = ( " " * $IndentFirst ) + $arrInputL[0]

	If( $arrInputL[0].Length -gt $intTextWidthL )
	{	#First "word" is longer than first line.
		#Chop it off and assign the rest to $strExtraL
		$strTempL = $arrInputL[ 0 ].SubString( 0, $intTextWidthL )
		$arrOutputL = @( $strTempL )
		$strExtraL = $arrInputL[0].SubString( $intTextWidthL )
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
			{	$arrOutputL += $strLMarginPadL + $strExtraL
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
					}
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
		{ Return $arrOutputL }
	Else	#Return a string
		{
		$strReturnL = ""
		ForEach( $strItemL in $arrOutputL )
			{ $strReturnL += "$strItemL`n" }
		Return $strReturnL
		}#End Else
}#End Paragraph #############################################################


##############################################################################
Function PrepareCredentialAd1_()
{#############################################################################
<#Examines script-scope variables and creates a psCredential object for
    Active Directory access.
  Typical use is to store an encrypted password in a file so script may be run
    without need to reenter password each time.
  Encrypted password is specific to the user account that created it.
  If stored in a shared location, like C:\ or C:\Temp, each user running the
    script will need to recreate the password file the first time they use it
    after another user has run the script and created their password file.
  If the account name, script argument, is blank, the user is prompted to supply
    a password, and current user account is used.
#>
	WriteVerbose_ 'Credential processing for Active Directory access...'
	If( $Script:RunAsUserAD -eq '' )
	{	$Script:RunAsUserAD = "$env:UserDomain\$env:UserName"
		WriteVerbose_ 'Using current user credential'
		WriteVerbose_ "  $Script:RunAsUserAD"
	}#End If - no user specified
	Else { WriteVerbose_ "  Credential account: $Script:RunAsUserAD" }

	#Default behavior - UsePwdFile

	If(	$Script:RunAsUsePwdFileAD     -eq $False	`
		-AND								`
		$Script:RunAsMakePwdFileAD    -eq $False	`
		-AND								`
		$Script:RunAsMakeNewPwdFileAD -eq $False	`
		)#End If
	{	$Script:RunAsUsePwdFileAD = $True }

	If( $Script:RunAsUsePwdFileAD )
	{	WriteVerbose_ '  Use existing password file.  Abort if file not found.'
		#If file of correct name is found, it is used.
		#If it contains the wrong password, that will be an error later...
		If( Test-Path $Script:RunAsPwdFileAD )
		{	WriteVerbose_ '  Existing file found.'
			$strPwdL =	Get-Content $Script:RunAsPwdFileAD	`
						|									`
						ConvertTo-SecureString
		}#End If - file exists
		#
		Else	#File not found
		{	$strMsgL =	'**Error - Encrypted password file not found for '	+	`
						'accessing Local Active Directory.'					+	`
						"`r`n"												+	`
						'  Run again with "-RunAsMakePwdFileAD" '			+	`
						'    or "-RunAsMakeNewPwdFileAD" switch.'
			WriteBoth_ $strMsgL -Red
			$Script:intReturnCodeG ++
			ExitProcessing_
		}#End Else - Pwd file not found
	}#End If - UsePwdFile
	#
	Else
	{	#One of two make new file options selected.
		#Check for Make File options...
		If( $Script:RunAsMakeNewPwdFileAD )
		{	WriteVerbose_ '  Create new file unconditionally.'
			$strMsgL = 'Enter password for local AD admin account'
			$strPwdL = Read-Host $strMsgL -AsSecureString
			WriteVerbose_ '  New password accepted.'
			$strPwdL	| ConvertFrom-SecureString			`
						| Out-File $Script:RunAsPwdFileAD
			WriteVerbose_ '  New password file created'
		}#End If - MakeNewPwdFile
		Else
		{	WriteVerbose_ '  Using existing password file if it exists.'
			If( Test-Path $Script:RunAsPwdFileAD )
			{	WriteVerbose_ '  Found existing file.'
				$strPwdL = Get-Content $Script:RunAsPwdFileAD	`
							| ConvertTo-SecureString
			}#End If - File exists
			#
			Else
			{	WriteVerbose_ "  File doesn't exist, create it."
				WriteBoth_ '  Creating encrypted password file...'
				$strPwdL = Read-Host ”Enter password:” -AsSecureString
				$strPwdL	| ConvertFrom-SecureString		`
							| Out-File $Script:RunAsPwdFileAD
			}#End Else - ! MakeNewFile
		}#End Else - Use if exists
	}#End Else - NOT UsePwdFile - MakeNewFile

	$Script:oCredentialAD =	New-Object											`
							-typename System.Management.Automation.PSCredential	`
								( $Script:RunAsUserAD, $strPwdL )
	WriteVerbose_ '  Local AD admin credential object created.'
}#End Prepare Credential AD ##################################################


##############################################################################
Function PrepareCredentialAd2_()
{#############################################################################
	WriteVerbose_ 'Credential processing for CITS Active Directory access...'
	If( $Script:RunAsUserAdTarget -eq '' )
	{	$Script:RunAsUserAdTarget = "$env:UserDomain\$env:UserName"
		WriteVerbose_ 'Using current user credential'
		WriteVerbose_ "  $Script:RunAsUserAdTarget"
	}#End If - no user specified
	Else { WriteVerbose_ "  Credential account: $Script:RunAsUserAdTarget" }

	#Default behavior - UsePwdFile

	If(	$Script:RunAsUsePwdFileAD      -eq $False	`
		-AND								`
		$Script:RunAsMakePwdFileAdT    -eq $False	`
		-AND								`
		$Script:RunAsMakeNewPwdFileAdT -eq $False	`
		)#End If
	{	$Script:RunAsUsePwdFileAdT = $True }

	If( $Script:RunAsUsePwdFileAdT )
	{	WriteVerbose_ '  Use existing password file.  Abort if file not found.'
		#If file of correct name is found, it is used.
		#If it contains the wrong password, that will be an error later...
		If( Test-Path $Script:RunAsPwdFileAdTarget )
		{	WriteVerbose_ '  Existing file found.'
			$strPwdL =	Get-Content $Script:RunAsPwdFileAdTarget	`
						|											`
						ConvertTo-SecureString
		}#End If - file exists
		#
		Else	#File not found
		{	$strMsgL =	'**Error - Encrypted password file not found for '	+	`
						'accessing Local Active Directory.'					+	`
						"`r`n"												+	`
						'  Run again with "-RunAsMakePwdFileAdT" '			+	`
						'    or "-RunAsMakeNewPwdFileAdT" switch.'
			WriteBoth_ $strMsgL -Red
			$Script:intReturnCodeG ++
			ExitProcessing_
		}#End Else - Pwd file not found
	}#End If - UsePwdFile
	#
	Else
	{	#One of two make new file options selected.
		#Check for Make File options...
		If( $Script:RunAsMakeNewPwdFileAdT )
		{	WriteVerbose_ '  Create new file unconditionally.'
			$strMsgL = 'Enter password for remote AD admin account'
			$strPwdL = Read-Host $strMsgL -AsSecureString
			WriteVerbose_ '  New password accepted.'
			$strPwdL	| ConvertFrom-SecureString			`
						| Out-File $Script:RunAsPwdFileAdTarget
			WriteVerbose_ '  New password file created'
		}#End If - MakeNewPwdFile
		Else
		{	WriteVerbose_ '  Using existing password file if it exists.'
			If( Test-Path $Script:RunAsPwdFileAdTarget )
			{	WriteVerbose_ '  Found existing file.'
				$strPwdL = Get-Content $Script:RunAsPwdFileAdTarget	`
							| ConvertTo-SecureString
			}#End If - File exists
			#
			Else
			{	WriteVerbose_ "  File doesn't exist, create it."
				WriteBoth_ '  Creating encrypted password file...'
				$strPwdL = Read-Host ”Enter password:” -AsSecureString
				$strPwdL	| ConvertFrom-SecureString		`
							| Out-File $Script:RunAsPwdFileAdTarget
			}#End Else - ! MakeNewFile
		}#End Else - Use if exists
	}#End Else - NOT UsePwdFile - MakeNewFile

	$Script:oCredentialADTarget =	New-Object											`
							-typename System.Management.Automation.PSCredential	`
								( $Script:RunAsUserAdTarget, $strPwdL )
	WriteVerbose_ '  Remote AD admin credential object created.'
}#End Prepare Credential AD Target ###########################################


##############################################################################
Function PrepareCredentialExLocal_()
{#############################################################################
<#Examines script-scope variables and creates a psCredential object for
    Exchange Server access in local domain.
  Typical use is to store an encrypted password in a file so script may be run
    without need to reenter password each time.
  Encrypted password is specific to the user account that created it.
  If stored in a shared location, like C:\ or C:\Temp, each user running the
    script will need to recreate the password file the first time they use it
    after another user has run the script and created their password file.
  If the account name, script argument, is blank, the user is prompted to supply
    a password, and current user account is used.
  Default should include $env:AppData as path component.
#>
	WriteVerbose_ 'Credential processing for local Exchange Server access...'
	If( $Script:RunAsUserExLocal -eq '' )
	{	$Script:RunAsUserExLocal = "$env:UserDomain\$env:UserName"
		WriteVerbose_ '  Using current user credential'
	}#End If - no user specified
	Else
	{	$strMsgL = "  Credential account: $Script:RunAsUserExLocal"
		WriteVerbose_ $strMsgL
	}#End Else

	#Default behavior - UsePwdFile

	If(	$Script:RunAsUsePwdFileExLocal     -eq $False	`
		-AND								`
		$Script:RunAsMakePwdFileExLocal    -eq $False	`
		-AND								`
		$Script:RunAsMakeNewPwdFileExLocal -eq $False	`
		)#End If
	{	$Script:RunAsUsePwdFileExLocal = $True }

	If( $Script:RunAsUsePwdFileExLocal )
	{	WriteVerbose_ '  Use existing password file.  Abort if file not found.'
		#If file of correct name is found, it is used.
		#If it contains the wrong password, that will be an error later...
		If( Test-Path $Script:RunAsPwdFileExLocal )
		{	WriteVerbose_ '  Existing file found.'
			$strPwdL =	Get-Content $Script:RunAsPwdFileExLocal	`
						|										`
						ConvertTo-SecureString
		}#End If - file exists
		#
		Else	#File not found
		{	$strMsgL =	'**Error - Encrypted password file not found for '	+	`
						'local Exchange Server access.'						+	`
						"`r`n"												+	`
						'  Run again with "-RunAsMakePwdFileExLocal" '		+	`
						'or "-RunAsMakeNewPwdFileExLocal" switch.'
			WriteBoth_ $strMsgL -Red
			$Script:intReturnCodeG ++
			ExitProcessing_
		}#End Else - Pwd file not found
	}#End If - UsePwdFile
	#
	Else
	{	#One of two make new file options selected.
		#Check for Make File options...
		If( $Script:RunAsMakeNewPwdFileExLocal )
		{	WriteVerbose_ '  Create new file unconditionally.'
			$strMsgL = 'Enter password for local Exchange admin account'
			$strPwdL = Read-Host $strMsgL -AsSecureString
			WriteVerbose_ '  New password accepted'
			$strPwdL	| ConvertFrom-SecureString				`
						| Out-File $Script:RunAsPwdFileExLocal
			WriteVerbose_ '  New password file created'
		}#End If - MakeNewPwdFile
		Else
		{	WriteVerbose_ '  Using existing password file if it exists.'
			If( Test-Path $Script:RunAsPwdFileExLocal )
			{	WriteVerbose_ '  Found existing file.'
				$strPwdL = Get-Content $Script:RunAsPwdFileExLocal	`
							| ConvertTo-SecureString
			}#End If - File exists
			#
			Else
			{	WriteVerbose_ "  File doesn't exist, create it."
				WriteBoth_ '  Creating encrypted password file...'
				$strMsgL = 'Enter password for local Exchange admin account'
				$strPwdL = Read-Host $strMsgL -AsSecureString
				$strPwdL	| ConvertFrom-SecureString		`
							| Out-File $Script:RunAsPwdFileExLocal
			}#End Else - ! MakeNewFile
		}#End Else - Use if exists
	}#End Else - NOT UsePwdFile - MakeNewFile

	$Script:oCredentialExLocal = New-Object										`
							-typename System.Management.Automation.PSCredential	`
										( $Script:RunAsUserExLocal, $strPwdL )
	WriteVerbose_ '  Local Exchange credential object created.'
}#End Prepare Credential ExLocal #############################################


##############################################################################
Function PrepareCredentialExRemote_()
{#############################################################################
<#Examines script-scope variables and creates a psCredential object for
    Exchange Server access in local domain.
  Typical use is to store an encrypted password in a file so script may be run
    without need to reenter password each time.
  Encrypted password is specific to the user account that created it.
  If stored in a shared location, like C:\ or C:\Temp, each user running the
    script will need to recreate the password file the first time they use it
    after another user has run the script and created their password file.
  If the account name, script argument, is blank, the user is prompted to supply
    a password, and current user account is used.
  Default should include $env:AppData as path component.
#>
	WriteVerbose_ 'Credential processing for remote Exchange Server access...'
	If( $Script:RunAsUserExRemote -eq '' )
	{	$Script:RunAsUserExRemote = "$env:UserDomain\$env:UserName"
		WriteVerbose_ '  Using current user credential'
	}#End If - no user specified
	Else
	{	$strMsgL = "  Credential account: $Script:RunAsUserExRemote"
		WriteVerbose_ $strMsgL
	}#End Else

	#Default behavior - UsePwdFile

	If(	$Script:RunAsUsePwdFileExRemote     -eq $False	`
		-AND								`
		$Script:RunAsMakePwdFileExRemote    -eq $False	`
		-AND								`
		$Script:RunAsMakeNewPwdFileExRemote -eq $False	`
		)#End If
	{	$Script:RunAsUsePwdFileExRemote = $True }

	If( $Script:RunAsUsePwdFileExRemote )
	{	WriteVerbose_ '  Use existing password file.  Abort if file not found.'
		#If file of correct name is found, it is used.
		#If it contains the wrong password, that will be an error later...
		If( Test-Path $Script:RunAsPwdFileExRemote )
		{	WriteVerbose_ '  Existing file found.'
			$strPwdL =	Get-Content $Script:RunAsPwdFileExRemote	`
						|											`
						ConvertTo-SecureString
		}#End If - file exists
		#
		Else	#File not found
		{	$strMsgL =	'**Error - Encrypted password file not found for '	+	`
						'remote Exchange Server access.'					+	`
						"`r`n"												+	`
						'  Run again with "-RunAsMakePwdFileExRemote" '		+	`
						'or "-RunAsMakeNewPwdFileExRemote" switch.'
			WriteBoth_ $strMsgL -Red
			$Script:intReturnCodeG ++
			ExitProcessing_
		}#End Else - Pwd file not found
	}#End If - UsePwdFile
	#
	Else
	{	#One of two make new file options selected.
		#Check for Make File options...
		If( $Script:RunAsMakeNewPwdFileExRemote )
		{	WriteVerbose_ '  Create new file unconditionally.'
			$strMsgL = 'Enter password for remote Exchange admin account'
			$strPwdL = Read-Host $strMsgL -AsSecureString
			WriteVerbose_ '  Accepted new remote Exchange password.'
			$strPwdL	| ConvertFrom-SecureString			`
						| Out-File $Script:RunAsPwdFileExRemote
			WriteVerbose_ '  New password file created'
		}#End If - MakeNewPwdFile
		Else
		{	WriteVerbose_ '  Using existing password file if it exists.'
			If( Test-Path $Script:RunAsPwdFileExRemote )
			{	WriteVerbose_ '  Found existing file.'
				$strPwdL = Get-Content $Script:RunAsPwdFileExRemote	`
							| ConvertTo-SecureString
			}#End If - File exists
			#
			Else
			{	WriteVerbose_ "  File doesn't exist, create it."
				WriteBoth_ '  Creating encrypted password file...'
				$strPwdL = Read-Host ”Enter password:” -AsSecureString
				$strPwdL	| ConvertFrom-SecureString		`
							| Out-File $Script:RunAsPwdFileExRemote
				WriteVerbose_ '  New password file created'
			}#End Else - ! MakeNewFile
		}#End Else - Use if exists
	}#End Else - NOT UsePwdFile - MakeNewFile

	$Script:oCredentialExRemote =	New-Object									`
							-typename System.Management.Automation.PSCredential	`
										( $Script:RunAsUserExRemote, $strPwdarrGroupListGL )
	WriteVerbose_ '  Remote Exchange credential object created.'
}#End Prepare Credential ExRemote ############################################


<#############################################################################
.Synopsis
Reads specified file, typically for configuration items.

.Description
Reads file and returns either an array of text lines or a hashtable of
  key/value pairs.
The -Hashtable option removes comment/blank lines then splits the active
  lines on the specified character, -Delimiter.
The first element from the line is trimmed and becomes the entry Key.
The remainining element(s) are not trimmed, and become the entry Value.
#>
Function ReadInputTextFile_()
{#############################################################################
Param(	#Fully qualified path to target file								`
		[Parameter( Mandatory = $True ) ]									`
		[string]$InputFile												,	`
		#Comment character used to ignore lines in input file				`
		[string]$CommentCharacter = '#'									,	`
		#Switch to return key/value pairs rather than text array			`
		[switch]$ReturnHashTable										,	`
		#Switch to allow continue if input file not Found					`
		[switch]$NoAbort												,	`
		#Delimiter used to create key/value pairs for hashtable option		`
		[string]$Delimiter        = ' '									,	`
		#Switch, remove whitespace from lines, if not hashtable				`
		[switch]$TrimLine													`
		)
	$colInputL = @()
	$strTargetFileL = $InputFile

	If( -NOT( Test-Path $strTargetFileL ) )
	{	#Not found.  Check script directory.
		$strTempPathL = EndBackslash $Script:strScriptPathG
		$strTargetFileL = $strTempPathL + $strTargetFileL
		If( -NOT( Test-Path $strTargetFileL ) )
		{	If( $NoAbort )
			{	If( $ReturnHashTable )
				{	$htReturnL = @{ Result = 'FileNotFound' }
					Return $htReturnL
				}#End If
				Else
				{	$colInputL = @( 'FileNotFound' )
					Return $colInputL
				}#End Else
			}#End If - NoAbort

			$strMsgL =	'**Error - Specified input file not found: '	+	`
						"`r`n`t$InputFile"								+	`
						"`r`n  Also not found as:"						+	`
						"`r`n`t$strTargetFileL"							+	`
						"`r`n  Exiting..."
			WriteBoth_ $strMsgL -Red
			ExitProcessing_
		}#End If - file not found in script folder either
	}#End If - file not found

	$arrInputL = Get-Content $strTargetFileL
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
	Else { 	Return $colTargetsL }
}#End Read Input Text File ###################################################


##############################################################################
Function ReadKeyValuePair_( $KVPair )
{#############################################################################
	#This only returns string values.
	$strPreferenceL = $KVPair.ToString()
	$strPreferenceL = $strPreferenceL.Replace( '[', '' )
	$strPreferenceL = $strPreferenceL.Replace( ']', '' )
	$arrReturnG     = $strPreferenceL.Split( ',' )
	Return $arrReturnG
}# End Read Key Value Pair ###################################################


##############################################################################
Function SendEmail_()
{#############################################################################
Param(	[int]$SmtpPort        = $Script:SmtpPort		,	`
		[string]$Attachments  = ''						,	`
		[string]$MsgBCC       = ''						,	`
		[string]$MsgBody      = ''						,	`
		[string]$MsgCC        = ''						,	`
		[string]$MsgFrom      = ''						,	`
		[string]$MsgTo        = ''						,	`
		[string]$SmtpAuthName = ''						,	`
		[string]$SmtpAuthPwd  = ''						,	`
		[string]$SmtpServer   = $Script:SmtpServer		,	`
		[string]$Subject      = ''						,	`
		[switch]$HTML									,	`
		[switch]$UseAuth									`
		)
	#To, CC and BCC are delimited lists with ; semicolon as delimiter
	If ( $MsgTo + $MsgCC + $MsgBCC -eq '' )
	{	WriteBoth_ '**Error - No recipients for email message.' -Red
		Return
	}#End If

	$strMsgL               =	$Script:strSeparatorG				+	`
								"  Sending email message... `r`n"
	$strErrorActionHoldL   = $ErrorActionPreference
	$ErrorActionPreference = 'Continue'
	Trap { ErrorHandler_ }

	$oMessageL      = New-Object System.Net.Mail.MailMessage
	$oMailFromL     = New-Object System.Net.Mail.MailAddress $MsgFrom
	$oMessageL.From = $oMailFromL
	$strMsgL       +=	"               Message from: $MsgFrom"	+	`
						"`r`n "							+	`
						'            Message format: '

	If( $HTML )
	{	$oMessageL.IsBodyHtml = $True
		$strMsgL += "HTML`r`n"
	}#End If
	Else { $strMsgL += "plain text`r`n" }

	#To Recipients
	If( $MsgTo -ne '' )
	{	$strMsgL += '                 Message To:'
		$arrTempL = $MsgTo.Split( ';' )
		ForEach( $strAddressL IN $arrTempL )
		{	$oMessageL.To.Add( $strAddressL )
			$strMsgL += "      $strAddressL`r`n"
		}#Next - recipient
	}#End If - To

	#CC Recipients
	If( $MsgCC -ne '' )
	{	$strMsgL += "                      CC To:`r`n"
		$arrTempL = $MsgCC.Split( ';' )
		ForEach( $strAddressL IN $arrTempL )
		{	$oMessageL.CC.Add( $strAddressL )
			$strMsgL += "      $strAddressL`r`n"
		}#Next
	}#End If - CC

	#BCC Recipients
	If( $MsgBCC -ne '' )
	{	$arrTempL = $MsgCC.Split( ';' )
		$strMsgL += "                     BCC To:`r`n"
		ForEach( $strAddressL IN $arrTempL )
		{	$oMessageL.BCC.Add( $strAddressL )
			$strMsgL += "      $strAddressL`r`n"
		}#Next
	}#End If - CC

	$oMessageL.Subject = $Subject
	$strMsgL += "                    Subject: $Subject`r`n"

	$oMessageL.Body    = $MsgBody

	If( $Attachments -ne '' )
	{	$strMsgL += "  Message has attachment(s):`r`n"
		$arrTempL = $Attachments.Split( ',' )
		ForEach( $strAttachmentL IN $arrTempL )
		{	If( Test-Path $strAttachmentL )
			{	$oAttachmentL = New-Object 										`
								System.Net.Mail.Attachment( $strAttachmentL )
			$oMessageL.Attachments.Add( $oAttachmentL )
			$strMsgL += "      $strAttachmentL`r`n"
		}#End If - attachment exists
		Else
		{	$strMsgL =	'**Error - Specified attachment file not found: '	+	`
						$strAttachmentL
			WriteBoth_ $strMsgL -Red
			ExitProcessing_
		}#End Else
  		}#Next - Attachment
	}#End - Attachments

	#Create the server session object
	$oSmtpClientL  = New-Object										`
						System.Net.Mail.SmtpClient	$SmtpServer	,	`
													$SmtpPort
	If( $UseAuth )
	{	$strMsgL += '  Using SMTP authentication.'
		$oCredentialsL = New-Object												`
							System.Net.NetworkCredential(	$SmtpAuthName	,	`
															$SmtpAuthPwd		`
														)
		$oSmtpClientL.Credentials = $oCredentialsL
	}#End If - Auth

	$strMsgL += "`r`n$strSeparatorG      Sending..."
	WriteVerbose_ $strMsgL

	$oSmtpClientL.Send( $oMessageL )
	If( $? )
	{	WriteVerbose_ '    Complete.' }
	Else { WriteVerbose_ '**Error - failed to send email...' }

	$oMessageL.Dispose()
	$ErrorActionPreference = $strErrorActionHoldL
}#End Send Email #############################################################


#############################################################################
Function SendEmailPS_()
{############################################################################
Param(	[string]$SmtpServer = $Script:oSmtpServerG.ServerName	,	`
		[int]$SmtpPort      = 25								,	`
		[string]$UserFrom										,	`
		[string]$Recipient										,	`
		[string]$BccTo											,	`
		[string]$CcTo											,	`
		[switch]$HTML											,	`
		[string]$Subject										,	`
		[string]$MsgBody										,	`
		[bool]$UseCred       = $Script:oSmtpServerG.UseAuth		,	`
		[switch]$UseSsl											,	`
		[string]$UserAuth    = $Script:oSmtpServerG.AuthName	,	`
		[string]$Password    = $Script:oSmtpServerG.AuthPwd		,	`
		[string]$Attachments										`
		)
	$blnReturnL   = $False
	$htParmsL     = @{}
	$strErrorL    = ''

	$htParmsL.Port = $SmtpPort
	If( $BccTo ) { $htParmsL.BCC    = $BccTo  }
	If( $CcTo )  { $htParmsL.CC     = $CcTo   }
	If( $HTML )  { $htParmsL.HTML   = $True   }
	If( $UseSsl ){ $htParmsL.UseSsl = $UseSsl }

	If( $Attachments )
	{	#This can be a string if only one attachment.
		#If multiple, must be array.
		#Can be array if only one.
		If( $Attachments -IS [Array] ){ $htParmsL.Attachments =	$Attachments }
		Else
		{	$arrTempL = $Attachments.Split( ',' )
			$htParmsL.Attachments =	$arrTempL
		}#End Else - not array
	}#End If - Attachments

	If( $SmtpServer ){ $htParmsL.SmtpServer = $SmtpServer }
	Else{ $strErrorL += "`r`n  **Error - SMTP Server required, none listed." }

	If( $UserFrom ){ $htParmsL.From = $UserFrom }
	Else{	$strErrorL += "`r`n  **Error - From address required, none listed." }

	If( $Recipient ){ $htParmsL.To = $Recipient }
	Else{	$strErrorL += "`r`n  **Error - Recipient required, none listed." }

	If( $Subject )
	{	$htParmsL.Subject = $Subject
	}#End If - Subject
	Else{	$strErrorL += "`r`n  **Error - Subject required, none listed." }

	If( $MsgBody )
	{	$htParmsL.Body = $MsgBody
	}#End If
	Else{	$strErrorL += "`r`n  **Error - Message body required, none listed." }

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
	Write-Verbose $strMsgL

	If( $strErrorL -ne '' )
	{	$strMsgL =	'  **The following problem(s) prevent '	+	`
					'sending this message.'					+	`
					"`r`n    $strErrorL"
		Write-Host $strMsgL -Foreground 'Red'
	}#End If - Error
	Else
	{	WriteVerbose_ '  Sending...'
		$htParmsL.EA = 'Continue'
		Send-MailMessage @htParmsL

		If( $? )
		{	$blnReturnL = $True
			WriteVerbose_ '    Completed successfully.'
		}#End If
		Else
		{	ErrorHandler_ 'Problem sending email.'
		}#End Else
	}#End Else - no error
	Return $blnReturnL
}#End Send Email PS ##########################################################


##############################################################################
Function SqlDateTime_()
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


##############################################################################
Function SqlEscapeQuotesInString_( [string]$InputString )
{#############################################################################
	$strOutputL = $InputString.Replace( "'", "''" )
	Return $strOutputL
}#End SQL Escape Quotes In String ############################################


##############################################################################
Function SqlExecuteQuery_( [string]$strQueryP )
{#############################################################################
	WriteVerbose_ "  Executing Query: $strQueryP"
	$oSqlCommandL = New-Object System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandText = $strQueryP
	$oSqlCommandL.Connection  = $Script:oDbConnectionG
	[Void]$oSqlCommandL.ExecuteNonQuery()
	[Void]$oSqlCommandL.Dispose()
}#End Function - SQL Execute Query ###########################################


##############################################################################
Function SqlInitializeDataset_( $strQueryP )
{#############################################################################
	$oSqlCommandL = New-Object System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandText    = $strQueryP
	$oSqlCommandL.Connection     = $Script:oDbConnectionG
	$oDataAdapterL               = New-Object System.Data.SqlClient.SqlDataAdapter
	$oDataAdapterL.SelectCommand = $oSqlCommandL

	WriteVerbose_ "  Executing query, $strQueryP"

	# Create and fill the DataSet object
	$oDataSetL = New-Object System.Data.DataSet
	$oDataAdapterL.Fill( $oDataSetL ) | Out-Null
	Return $oDataSetL
}#End Function - SQL Initialize Dataset ######################################


##############################################################################
Function SqlInitializeDbConnection_()
{#############################################################################
Param(	[string]$Server					,	`
		[string]$Database				,	`
		[string]$Port     = "1433"		,	`
		[string]$Instance = ""			,	`
		[string]$UserId   = "Windows"	,	`
		[string]$Pwd      = ""				`
		)
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
	WriteVerbose_ "  Executing DB Connection with string: $strConnectionL"

	$oDbConnectionL = New-Object System.Data.SqlClient.SqlConnection
	$oDbConnectionL.ConnectionString = $strConnectionL
	If( -NOT $? )
	{	$strErrL = "While establishing connection string: $strConnectionL"
		ErrorHandler_( $strErrL )
	}#End If - Error

	$oDbConnectionL.Open()
	If( -NOT $? )
	{	$strErrL =	"While opening connection with "	+	`
					"connection string: $strConnectionL"
		ErrorHandler_( $strErrL )
	}#End If - Error
	Return $oDbConnectionL
}#End SQL Initialize DB Connection ###########################################


##############################################################################
Function SqlInitializeReader_( $strQueryP )
{#############################################################################
	WriteVerbose_ "  Initializing dataReader with query: $strQueryP"
	$oSqlCommandL = New-Object System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandTimeout = 30
	$oSqlCommandL.Connection     = $Script:oDbConnectionG
	$oSqlCommandL.CommandText    = $strQueryP

	If( -NOT $? )
	{	$strErrL = "While creating new SQL Command object."
		ErrorHandler_( $strErrL )
	}#End If - Error

	$Script:oSqlReaderG = $oSqlCommandL.ExecuteReader()
	If( -NOT $? )
	{	$strErrL = "While executing SQL Reader with, $strQueryP"
		ErrorHandler_( $strErrL )
	}#End If - Error
}#End Function - SQL Initialize Reader #######################################


##############################################################################
Function SqlInsertRecord_()
{#############################################################################
Param(	[string]$TableName	,	`
		$DataTable				`
		)
#Accepts TableName and hash table.
#Hash table is column name (name) and data (value) for one record.
#Build lists of column names and values for query.
#Add .Parameters.AddWithValue, for each column, to command object.
#Then execute command.

	$strFieldL      = ''
	$strParmNameL   = ''
	$strTableL      = ''
	$arrFieldNamesL = $DataTable.Keys

	$oSqlCommandL            = New-Object System.Data.SqlClient.SqlCommand
	$oSqlCommandL.Connection = $Script:oDbConnectionG

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
	WriteVerbose_ "  Insert Query: $strQueryL"
	$oSqlCommandL.CommandText = $strQueryL

	$intRowsL = $oSqlCommandL.ExecuteNonQuery()
	[void]$oSqlCommandL.Dispose()

	If( $intRowsL -ne 1 )
	{	strMsgL =	"  **Error - While inserting record in table, "	+	`
					$strTableL
		WriteBoth_ $strMsgL
		ExitProcessing_
	}#End If - Error
}#End Function - SQL Insert Record ###########################################


##############################################################################
Function SqlRecordExists_()
{#############################################################################
Param(	[string]$Table, [string]$WhereClause )
#Checks to see if a record exists in a table, based on a key and value.
#Only works for strings, dates, etc. where value is quoted in query.
	$strQueryL =	'SELECT * FROM '	+	`
					$Table				+	`
					' '					+	`
					$WhereClause
	WriteVerbose_ $strQueryL

	$oSqlCommandL = New-Object System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandTimeout = 30
	$oSqlCommandL.Connection     = $Script:oDbConnectionG
	$oSqlCommandL.CommandText    = $strQueryL

	If( -NOT $? )
	{	$strErrL = "While creating new SQL Command object."
		ErrorHandler_( $strErrL )
	}#End If - Error

	$oSqlReaderL = $oSqlCommandL.ExecuteReader()
	If( -NOT $? )
	{	$strErrL = "While executing SQL Rearder with, $strQueryL"
		ErrorHandler( $strErrL )
	}#End If - Error

	$blnFoundL = $oSqlReaderL.HasRows
	[void]$oSqlReaderL.Close()
	Return $blnFoundL
}#End - SQL Record Exists  ###################################################


##############################################################################
Function SqlTextDelim_( [string]$InputText )
{#############################################################################
	$strTempL = $InputText.Replace( "'", "`'`'" )
	####$strTempL = $InputText.Replace( "\", ""'\"" )
	Return $strTempL
}#End SQL Text Delim #########################################################


##############################################################################
Function SqlTruncateTable_( $TableToTruncate )
{#############################################################################
	$strQueryL = "DELETE $TableToTruncate"
	$oSqlCommandL = New-Object System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandText = $strQueryL
	$oSqlCommandL.Connection  = $Script:oDbConnectionG
	[Void]$oSqlCommandL.ExecuteNonQuery()
	[Void]$oSqlCommandL.Dispose()
}#End Function - SQL Truncate Table #########################################


##############################################################################
Function SqlUpdateRecords_()
{#############################################################################
Param(	[Parameter( Mandatory = $True, HelpMessage = 'TableName required' ) ]		`
		[string]$TableName														,	`
		[Parameter( Mandatory = $True, HelpMessage = 'WHERE required' ) ]			`
		[string]$WhereClause													,	`
		[Parameter( Mandatory = $True, HelpMessage = 'HT required' ) ]				`
		$DataTable																,	`
		[Switch]$Test															,	`
		[Switch]$Create																`
		)
#Typical call:
#	SqlUpdateRecords_	-TableName   'TableName'	`
#						-WhereClause $strQueryL		`
#						-DataTable   $htDataL		`
#						-Test						`
#						-Create
#Accepts TableName, Where clause and hash table
#DataTable is hash table with column names as keys (name) and data values,
#  for the fields whose values will be updated.
#-Test confirms existence of record before attempting update.
#-Create inserts record if -Test doesn't find it.
	If( $Test )
	{	If( -NOT( SqlRecordExists_ -Table $TableName -WhereClause $WhereClause ) )
		{	If( $Create )
			{	$strMsgL =	'  Record requested for Update does not exist.'	+	`
							"`r`n"											+	`
							'  Create option specified.'					+	`
							"`r`n"											+	`
							'  Record will be inserted.'
				WriteVerbose_ $strMsgL
				SqlInsertRecord_ -TableName $TableName -DataTable $DataTable
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
				WriteBoth_ $strMsgL -Red
				ExitProcessing_
			}#End Else - no create
		}#End If - record doesn't exist
		####}#End Else - record not found
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
	WriteVerbose_ $strQueryL

	$oSqlCommandL = New-Object System.Data.SqlClient.SqlCommand
	$oSqlCommandL.CommandText = $strQueryL
	$oSqlCommandL.Connection  = $Script:oDbConnectionG
	$intRowsL = $oSqlCommandL.ExecuteNonQuery()
	If( $intRowsL -ne 1 )
	{	$strMsgL =	"  **Error - While updating record(s) "	+	`
					"in table, "							+	`
					$strTableL
		WriteBoth_( $strMsgL )
	}#End If - <> 1
	[Void]$oSqlCommandL.Dispose
}#End Function - SQL Update Records ##########################################


##############################################################################
Function SleepTimer_( [string]$Message, [int]$Duration )
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
	###Write-Progress -Completed
}#End Sleep Timer ############################################################


##############################################################################
Function WriteEventLogEntry_()
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
	If( ( $Source -eq $Null ) -OR ( $Source -eq '' ) )
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


<#############################################################################
.Synopsis
Returns boolean indicating whether current time is within bounds of parms.

.Description
Start/Stop can be dateTime objects or strings.
Strings can be 1-2 digit integers representing hours in 24 hour format, or
  strings representing date and or time.
Any portion of a date time string can be used if it can be cast as DateTime.
For instance, [dateTime]'8/20/2014' returns a dateTime object for 12:00 AM
  on 8/20/2014.
Converting time, such as '11:30', will return as 11:30 AM on current date.
#>
Function WithinRuntime_()
{#############################################################################
Param(	$Start				,	`
		$Stop				,	`
		[switch]$Recurring		`
		)
	If( $Script:Continuous ){ Return $True }

	$blnIntegerL  = $False
	$intHourNowL  = ( Get-Date ).Hour
	$oTypeL       = $Start.GetType()
	$strTimeTypeL = $oTypeL.Name
	$strMsgL      = "  Start parameter type: $strTimeTypeL"
	WriteVerbose_ $strMsgL

	If( $strTimeTypeL -eq 'String' )
	{	$arrTempL = GetTime_ $Start
		If( $arrTempL[0] -eq 'Integer' )
		{	$strMsgL     = '  String type argument in integer value.'
			WriteVerbose_  $strMsgL
			$blnIntegerL = $True
			$intStartL   = $arrTempL[1]
			$arrTempL    = GetTime_ $Stop
			$intStopL    = $arrTempL[1]
			$strMsgL     = '  Start integer: '	+	`
							$intStartL			+	`
							', Stop: '			+	`
							$intStopL			+	`
							', Now: '			+	`
							$intHourNowL
			WriteVerbose_ $strMsgL

			If( $intStartL -ge 12 )	#Start time is PM
			{	If( $intStopL -ge 12 )	#Stop also PM
				{	If( $intHourNowL -ge $intStartL	`
						-AND						`
						$intHourNowL -lt $intStopL	`
						)
					{	Return $True }
					Else { Return $False }
				}#End If - stop is PM ( so is start )
				Else	#Start is PM, stop time is in AM
				{	If( $intHourNowL -ge 12 )	#now is PM
					{	If( $intHourNowL -ge $intStartL ){ Return $True }
						Else { Return $False }
					}#End If - current hour PM ( stop time AM )
					Else	#current hour AM ( stop time AM )
					{	If( $intHourNowL -lt $intStopL ){ Return $True }
						Else { Return $False }
					}#End Else - current hour AM ( stop time AM )
				}#End Else - stop AM
			}#End If - start PM
			Else	#Start time is AM
			{	If( ( $intHourNowL -ge $intStartL )	`
					-AND							`
					( $intHourNowL -lt $intStopL )	`
					)
				{	Return $True }
				Else { Return $False }
			}#End Else - start is AM
		}#End If - integer
		Else #argument was string, but not integer.
		{	$strTimeTypeL = 'DateTime'
			$arrTempL     = GetTime_ $Start
			$dteStartL    = $arrTempL[1]
			$arrTempL     = GetTime_ $Stop
			$dteStopL     = $arrTempL[1]
			$strMsgL      = '  Start time: '		+	`
							$dteStartL				+	`
							"`r`n   Stop time: "	+	`
							$dteStopL				+	`
							"`r`n         Now: "	+	`
							( Get-Date ).ToString()
			WriteVerbose_ $strMsgL
		}#End Else - string, not integer
	}#End If - string
	Else	#DateTime as argument
	{	$blnDateTimeL = $True
		$arrTempL     = GetTime_ $Start
		$dteStartL    = $arrTempL[1]
		$arrTempL     = GetTime_ $Stop
		$dteStopL     = $arrTempL[1]
	}#End Else - dateTime

	If( $Recurring )
	{	#Recurring, continuous run multiple days.
		#Arguments are integers representing hours in 24 hour format.
		If( -NOT( $blnIntegerL ) )
		{	$strMsgL =	'**Error - Invalid argument combination'			+	`
						"`r`n  When specifying the recurring option, "		+	`
						'specify start/stop as integers, 24 hour format.'
			WriteBoth_ $strMsgL
			Exit
		}#End If - not integer

		If( $intStartL -ge 12 )
		{	#Start time is PM
			If( $intStopL -ge 12 )
			{	If( ( Get-Date ) -ge $intStopL	`
					-AND						`
					( Get-Date ) -lt $Stop		`
					)
				{	Return $True }
				Else { Return $False }
			}#End If - stop PM
			Else
			{	#Stop time is in AM
				If( $True ) {}
			}#End Else - stop AM
		}#End If - start PM
	}#End If - Recurring
	Else
	{	#Not a recurring run, single start/stop time.
		If( $blnIntegerL )
		{	$intHourNowL = ( Get-Date ).Hour
			If(	( $intStartL -ge $intHourNowL )	`
				-AND							`
				( $intHourNowL -lt $intStopL )	`
				)
		{	Return $True }
		Else { Return $False }
		}#End If - integer
		Else	#using dateTime values not integers
		{	If( ( Get-Date ) -ge $intStopL	`
			-AND							`
			( Get-Date ) -lt $stop			`
			)
		{	Return $True }
		Else { Return $False }
		}#End Else - not integer
	}#End Else - not recurring
}# End Within Runtime ########################################################


##############################################################################
Function WriteBanner_()
{#############################################################################
	$strMsgL =	"`r`n`r`n"				+	`
				$Script:strSeparatorG	+	`
				'  '					+	`
				$Script:VersionInfo		+	`
				"`r`n    "				+	`
				( Get-Date )			+	`
				"`r`n"					+	`
				$Script:strSeparatorG
	WriteBoth_ $strMsgL
}#End Write Banner ###########################################################


##############################################################################
Function WriteBoth_( [string]$InputP, [string]$Color = $Null, [switch]$Red )
{#############################################################################
	If( $Red )
	{	$Color = 'Red'
		Write-Host $InputP -ForegroundColor $Color
	}#End If - Red
	Else
	{	If( $Color )
		{	$arrTempL = $Color.Split( '|' )
			If( $arrTempL.Length -gt 1 )
			{	#Color combo w/ background specified
				$Color    = $arrTempL[0]
				$strBackL = $arrTempL[1]
				Write-Host	-Object $InputP				`
							-ForegroundColor $Color		`
							-BackgroundColor $strBackL
			}#End If - two colors
			Else{ Write-Host $InputP -ForegroundColor $Color }
		}#End - Color
		Else{ Write-Host $InputP }
	}#End Else - not Red
	WriteLog_ $InputP
}#End Write Both #############################################################


##############################################################################
Function WriteLog_( $InputP )
{#############################################################################
	If( $Script:blnLogInitializedG )
	{	$InputP | Out-File	-FilePath $Script:strLogFileFqG	`
							-Encoding "Default"				`
							-Append
	}#End If
	Else
	{	$Script:arrLogHoldG += $InputP }
}#End Write Log ##############################################################


##############################################################################
Function WriteVerbose_()
{#############################################################################
Param(	[string]$Message	,	`
		[switch]$NoHost		,	`
		[switch]$NoLog			`
		)
	$strEaL = $ErrorActionPreference	#Compatibility after introduce Details
	$ErrorActionPreference = 'SilentlyContinue'
	If( ( $VerbosePreference -eq 'Continue' ) -OR ( $Script:Details ) )
	{	$Message = "<V> $Message"
		If( -NOT( $NoHost ) ){ Write-Host $Message -ForegroundColor Blue }
		If( -NOT( $NoLog  ) ){ WriteLog_  $Message }
	}#End If
	$ErrorActionPreference = $strEaL
}#End Write Verbose ##########################################################

#Main
#Write-Host "This is the Library file.  It doesn`'t do anything by itself..."
#If( $Script:ShowVersion )
#{	CLS
#	$strVersionG = GetLibraryVersion_
#	Write-Host $strVersionG
#	Exit
#}#End If