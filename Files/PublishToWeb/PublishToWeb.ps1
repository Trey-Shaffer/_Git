Param(	#See general comments at end of file.									`
		#List of Scripts to be processed.										`
		#Individual, comma-separated or file name holding list.					`
		#Values should be fully qualified file names of scripts.				`
		#Empty ('') string value will retrieve all scripts from database.		`
		[string]$Scripts = ''												,	`
		#Default location to publish scripts.									`
		[string]$PublishFolder =	'D:\Documents\Visual Studio\repos\'		+	`
									'ITArtisan\ITArtisan\z_ScriptContent'	,	`
		#Overwrite existing target script and hash files.						`
		[bool]$OverWrite = $True											,	`
		#Use PublishFolder values from individual database records.				`
		#	If DB field empty, still use default, $PublishFolder.				`
		[bool]$UseDbPublishPath = $False									,	`
		#Switch to disable update of database.									`
		[switch]$noDbUpdate													,	`
		#Display string for banner												`
		[string]$VersionInfo = 'Publish Scripts, v1.0 - July 2017'			,	`
		#Version number detailed.												`
		[string]$VersionNumber = '1.0'										,	`
		#Root for log file name, with possible date/time addition				`
		[string]$LogBase = 'Pub'											,	`
		#Path to folder holding script log file.  '' = script folder			`
		[string]$LogFolder = ''												,	`
		#
		#Values associated with notification email(s).							`
		#Specify a fully qualified file name for attachment(s).					`
		#Comma-delimited list, or array, of attachments allowed.				`
		#Empty, '', value will not cause an attachment.							`
		#A value of '*IncludeLogFile*' will include log file from this script.	`
		[string]$AlertAttachment   = '*IncludeLogFile*'						,	`
		#All three recipient values may be Comma-delimited lists.				`
		[string]$AlertRecipient    = 'Trey.Shaffer@Charter.com'				,	`
		[string]$AlertRecipientBCC = ''										,	`
		[string]$AlertRecipientCC  = ''										,	`
		[string]$AlertFrom         = 'PowerShell@Charter.com'				,	`
		[string]$AlertSubject      = '**Alert - '							,	`
		#Switch to cast email in HTML format.									`
		[switch]$AlertHTML													,	`
		#Switch disables sending of email alerts.								`
		[switch]$NoEmail													,	`
		#																		`
		#Secondary/common parameters											`
		#File name only of ps1 file to include									`
		[string]$ModuleFile = 'PsLibrary2'									,	`
		#Path to library module, for use if not available locally				`
		#Production																`
			[string]$Modulepath = '\\OpMgr02\PSModules$'					,	`
		#Lab																	`
			#\\LabOpMgr01\PSModules$'										,	`
		#Minimum acceptable version of library File								`
		[string]$ModuleMinVersion = '1.0'									,	`
		#Default delimiter to use in lists										`
		[string]$Delimiter = ','											,	`
		#Host name of SQL Server												`
		#Production																`
			[string]$SqlServer = 'SQL17-01'									,	`
		#Lab																	`
			#[string]$SqlServer = 'Lab-SqlClus01'							,	`
		#SQL Server database name												`
		[string]$SqlDatabase = 'ScriptRepo	'								,	`
		#SQL Server TCP port used (default=1433)								`
		[int]$SqlPort = 1433												,	`
		#SQL Server instance name, if used										`
		[string]$SqlInstance = ""											,	`
		#SQL Server user name.  Use "Windows" for integrated security.			`
		[string]$SqlUser = "Windows"										,	`
		#SQL Server password if not using integrated security					`
		[string]$SqlPwd = ""												,	`
		#																		`
		#Extension for log file name.  Should include leading dot "."			`
		[string]$LogExt = '.txt'											,	`
		#Switch, append to existing log file, rather than overwrite.			`
		[switch]$LogAppend													,	`
		#Switch, Add date string to log file name								`
		[switch]$LogDate													,	`
		#Switch, Add time string to log file name								`
		[switch]$LogTime													,	`
		#Switch, enable verbose output for script								`
		[switch]$Verbose2														`
		)
##############################################################################
#Seldom used parameters
#File name only of Access database										`
#[string]$MdbFile = 'SomeDatabase.mdb'								,	`
#Path to folder holding Access database									`
#[string]$MdbFolder = ''											,	`
##############################################################################
Function EndBackslash( [string]$strInput )
{#############################################################################
	If(($strInput -eq '' ) -OR ($strInput.EndsWith('\'))){ Return $strInput }
	Else{ Return "$strInput\" }
}#End End Backslash ##########################################################


##############################################################################
Function GetScriptPath(){ Return Split-Path $MyInvocation.ScriptName -parent }
#There's a reason this only works in a function, but I forget it...
##############################################################################


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
Function LoadModuleLocal()
{#############################################################################
Param(	#Name of module to loaded									`
		[string]$ModuleName										,	`
		#Path to source holding module to install and load			`
		[string]$ModulePath = $Script:ModulePath				,	`
		#Minimum version for ModuleFile								`
		[int]$ModuleMinVersion = '1.0'							,	`
		#Switch to only check load status no attempt to load		`
		[switch]$CheckOnly										,	`
		#Switch to allow continue after failure to load				`
		[switch]$NoAbort											`
		)
	$strMsgL = "  Examining status of module: $ModuleName..."
	Write-Verbose $strMsgL

	If( Get-Module -Name $ModuleName ){ Return $True }
	Else
	{	#Did not find it in loaded modules.
		If( $CheckOnly ){ Return $False }	#Not attempting load

		Write-Host '  Checking locally available modules...'
		If( Get-Module -ListAvailable |	`
				Where-Object{$_.Name -eq $ModuleName }
			)
		{	#Not loaded, but available, try to load it
			$strEaHoldL            = $ErrorActionPreference
			$ErrorActionPreference = 'SilentlyContinue'
			Write-Verbose '  Loading module...'
			Import-Module $ModuleName

			If( $? ){ Return $True }
			Else
			{	If( $NoAbort )
				{	$ErrorActionPreference = $strEaHoldL
					Return $False }
				Else{ Exit 77 }
			}#End If - error
		}#End If - module available
	}#End Else - Not loaded
	#At this point module is not loaded or locally avaialable.

	If( $ModulePath -eq $Null )
	{	#No source path was specified to install module.
		$strMsgL =	'**Warning - specified module not '	+	`
					'loaded OR available locally.'		+	`
					"`r`n  Alternate installation "		+	`
					'path not specified.'
		Write-Host $strMsgL -ForegroundColor 'Yellow'
		Return $False
	}#End If - no ModulePath
	Else
	{	If( Test-Path $ModulePath )
		{	Write-Verbose '  Source module path confirmed.'
			$strSourceFolderL  = EndBackSlash $ModulePath
			$strSourceFolderL += $ModuleName
		}#End If
		Else
		{	$strMsgL =	'**Error - Module source folder not found.'	+	`
						"`r`n  -ModulePath, $ModulePath"
			Write-Host $strMsgL -ForegroundColor 'Red'
		}#End Else
	}#End Else - Module path supplied

	Write-Verbose "`r`n  Module is not available locally."
	#Now look at profile path, then \Modules to see if exists
	$strProfileL       = $Profile
	$strProfileFolderL = Split-Path $strProfileL
	$strMsgL           = "`r`n  Local profile folder:`r`n`t$strProfileFolderL"
	Write-Verbose $strMsgL

	#Modules folder may not exist.
	#Not created until first module installed/loaded
	$strModulesPathL = "$strProfileFolderL\Modules\"

	#Check for Modules Folder
	If( Test-Path $strModulesPathL )
	{	Write-Verbose "`r`n  Local Modules folder exists" }
	Else
	{	Write-Verbose "`r`n  Modules folder does not exist.  Creating it..."
		$oModulesPathL = New-Item $strModulesPathL -Type Directory
	}#End If - no modules folder

	#Target folder for specified module.
	$strTargetFolderL = $strModulesPathL + $ModuleName

	#Check for target module folder
	If( Test-Path $strTargetFolderL )
	{	$strMsgL =	"`r`n  Local Target module folder exists: "	+	`
					"`r`n`t$strTargetFolderL"
		Write-Verbose $strMsgL
	}#End If - Test-Path
	Else
	{	#Target folder for module not found in profile folder.
		$strMsgL =	"`r`n  Local target module folder does not "	+	`
					"exist: `r`n`t"									+	`
					$strTargetFolderL								+	`
					"`r`n  Will copy from $strSourceFolderL"
		Write-Verbose $strMsgL
	}#End Else - Target folder not found.

	Copy-Item	-Path        $strSourceFolderL	`
				-Destination $strModulesPathL	`
				-EA          'SilentlyContinue'	`
				-Recurse						`
				-Force

	#Module now available, load it.
	$strEaHoldL            = $ErrorActionPreference
	$ErrorActionPreference = 'SilentlyContinue'
	Write-Verbose '  Loading module...'
	Import-Module $ModuleName

	If( !$? )
	{	#ErrorHandler_ "Failed to load module, $ModuleName"
		If( $NoAbort )
		{	$ErrorActionPreference = $strEaHoldL
			Return $False
		}#End If
		Else{ Exit 77 }
	}#End If - error
	$ErrorActionPreference = $strEaHoldL
	Return $True
}# End Local Load Module #####################################################


<#############################################################################
.Synopsis
Processes one script record.  Copies from dev location to web Publish
	location and updates file hash value and file.

.Description
Accept a database record holding metadata about a script file.
Copy file from source (DevPath) to target (PublishPath).
Create a related file with the specified hash value and update database
	with hash value.
#############################################################################>
Function ProcessRecord()
{#############################################################################
[CmdletBinding()]
[OutputType( "System." )]
Param(	$Record												,	`
		[string]$PublishFolder  = $Script:PublishFolder		,	`
		[bool]$OverWrite        = $Script:OverWrite			,	`
		[bool]$UseDbPublishPath = $Script:UseDbPublishPath	,	`
		[bool]$noDbUpdate       = $Script:NoDbUpdate			`
		)
	$strFileNameL      = $Record.FileName
	$strScriptNameL    = $Record.ScriptName
	$strFileHashL      = $Record.FileHash
	$strPublishFolderL = $Record.PublishFolder
	
	If( $UseDbPublishPath = $False )
	{	$strPublishPathL = $PublishFolder
		
		#If not specified, use default publish path.
		If( [string]$strPublishFolderL -EQ '' )
		{	$strPublishFolderL = $PublishFolder }
	}#End If - UseDbPublishPath
	Else { $strPublishFolderL = $PublishFolder }
	
	#Verify PublishPath
	If( Test-Path $strPublishFolderL )
	{	Write-Verbose_ 'Verified publish path.'}
	Else
	{	$strMsgL =	'**Error - Publish folder not found'		+	`
					"`r`n  Failed to locate: $oTargetFolderL"
		Write-Both_ -InputText $strMsgL -Red
		Return $False
	}#End If - target publish folder not found
	
	#Verify source file exists.
	$strFileNameL = $strFileNameL.Replace( '"', '')
	$oSourceFileL = Get-ChildItem -Path $strFileNameL
	If( $oSourceFileL -EQ $Null )
	{	$strMsgL =	'**Error - Source file not found'		+	`
					"`r`n  Failed to locate: $strFileNameL"
		Write-Both_ -InputText $strMsgL -Red
		Return $False
	}#End If - source not found
	
	#Construct FQ target file names.
	$strTargetFileNameFqL = "$strPublishFolderL$strScriptNameL"
	$strTargetHashFileFqL = "$strTargetFileNameFqL.MD5.txt"
	
	#Get target file to see if it exists.
	#$oTargetFileL = Get-ChildItem -LiteralPath $strTargetFileNameFqL
	
	If( Test-Path $strTargetFileNameFqL )
	{	#File exists, should it be deleted?
		If( $OverWrite )
		{	#Yes, delete target file before copy.
			
			#Delete target hash file.
			If( Test-Path $strTargetHashFileFqL )
			{ Remove-Item $strTargetHashFileFqL }
		}#End If - OverWrite
		Else
		{	#Not overwriting.  Do nothing.  No action for hash file.
			Return $True
		}#End Else
	}#End If - target file found
	
	#Target files either did not exist or have been deleted.
	
	#Copy source script to target.
	$oResultL = Copy-Item	-LiteralPath $strFileNameL			`
							-Destination $strTargetFileNameFqL	`
							-Force
	
	#Calculate hash value.
	$oHashL   = Get-FileHash $oSourceFileL -Algorithm 'MD5'
	$strHashL = $oHashL.Hash
	
	#Write hash to file.
	$strHashL | Out-File -FilePath $strTargetHashFileFqL
	
	#Update database with new hash value and last published date.
	$strDateL = ConvertTo-SqlDateTime_	-DateTime ( Get-Date )	`
										-AddQuotes
	#Update database
	If( $Script:NoDbUpdate -EQ $False )
	{	$htDataL               = @{}
		$htDataL.FileHash      = "'$strHashL'"
		$htDataL.LastPublished = $strDateL
		
		$strQueryL =	"WHERE FileName = '"	+	`
						$strFileNameL			+	`
						"'"
		
		Invoke-SqlUpdateRecords_	-TableName   'Scripts'	`
									-WhereClause $strQueryL	`
									-DataTable   $htDataL	`
									-Test
	}#End If - NoDbUpdate
	Return $True
}#End - Process Record #######################################################



##############################################################################
###Main
######################################################################
If( $True ){ #Common base code
CLS
Trap { New-ErrorInfo_ -Error $_ }

Set-PsDebug -Strict
$Global:dteStartG             = Get-date
$Global:blnLogInitializedG    = $False
$Script:ErrorActionPreference = 'Stop'
$Script:ErrorDescriptionG     = 'No additional information'
$Script:arrLogHoldG           = @( "First Line" )
$Script:intReturnCodeG        = 0
$Script:strScriptPathG        = GetScriptPath
$Script:strSeparatorG         = ( '=' * 78 ) + "`r`n"
$Script:strThisScriptG        = Split-Path -Path $Script:strScriptPathG -Leaf

##$Script:DebugMe = $True
#If( $Script:DebugMe )
#{	$strDebugOutG = "$Script:strScriptPathG\Debug.txt"
#	$strMsgG      = ( Get-Date ).ToString()
#	$strMsgG      | Out-File -FilePath $strDebugOutG -Encoding "ASCII"
#}#End If - debug

$Host.UI.RawUI.WindowTitle = "$Script:VersionInfo - $Script:strScriptPathG"

If( $Script:Verbose2 ){ $VerbosePreference = 'Continue' }
Else{ $VerbosePreference = 'SilentlyContinue' }

$blnLoadedG = LoadModuleLocal	-ModuleName       $Script:ModuleFile		`
								-ModulePath       $Script:ModulePath		`
								-ModuleMinVersion $Script:ModuleMinVersion
If( $blnLoadedG -eq $False ){ Exit 66 }
$strVersionLoadedL = Get-PSLibraryVersion_
$strMsgG           = "  Running PSLibrary.psm1 version $strVersionLoadedL"
Write-Verbose $strMsgG
If( $strVersionLoadedL -lt $Script:ModuleMinVersion )
{	$strMsgL =	'**Warning - Module loaded incorrect version...'	+	`
				"`r`n  Requested module, $Script:ModuleName "		+	`
				"version $strVersionLoadedL loaded"					+	`
				"`r`n  Version $Script:ModuleMinVersion required."
	Write-Host $strMsgL -Color 'Yellow|Black'
	Exit 78
}#End Else

#If needed, force date/time stamping here...
If( $Script:LogFolder -eq '' ){ $Script:LogFolder = $Script:strScriptPathG }
$Global:strLogFileFqG = Initialize-Log_	-LogAppend $Script:LogAppend	`
										-LogBase   $Script:LogBase		`
										-LogExt    $Script:LogExt		`
										-LogDate   $Script:LogDate		`
										-LogFolder $Script:LogFolder	`
										-LogTime   $Script:LogTime
If( $Global:strLogFileFqG ) { $Global:blnLogInitializedG = $True }

#Enable this section if using any database access.
#If getting script names from database...
If( $Script:Scripts -NE '' )
{	#Getting script names from database.  Initialize database.
	Write-Verbose_ '  Initializing database connection...'
	$Global:oDbConnectionG = New-SqlDbConnection_								`
												-Database $Script:SqlDatabase	`
												-Instance $Script:SqlInstance	`
												-Port     $Script:SqlPort		`
												-Pwd      $Script:SqlPwd		`
												-Server   $Script:SqlServer		`
												-UserID   $Script:SqlUser
}#End If - using database

#Enable this if using email.
#Email configuration requires database information.
#$Global:oSmtpServerG = Get-SmtpServer_ $Global:oDbConnectionG

If( $Script:LogAppend )
{	Write-Banner_	-VersionInfo $VersionInfo	`
					-Separator 	 $strSeparatorG	`
					-Append
}#End If - append
Else
{	Write-Banner_	-VersionInfo $VersionInfo	`
					-Separator 	 $strSeparatorG
}#End Else - no append

}#End common base code
##############################################################################
$Script:PublishFolder = Edit-EndChar_ $Script:PublishFolder
	
#Build list of script names to process, either from argument or database.
If( $Script:Scripts -EQ '' )
{ #Get script names from database
	Write-Verbose_ '  Getting script names from database...'
	$strQueryG = 'SELECT FileName FROM SCRIPTS'
	$oDataSetG = New-SqlDataset_ -Query $strQueryG
		
	#Build array of file names.
	$colFileNamesG = @()
	
	# Iterate through dataset
	ForEach( $oRowG IN $oDataSetG.Tables[0] )
	{	$colFileNamesG += $oRowG.FileName }
	
	If( $colFileNamesG.Count -EQ 0 )
	{	$strMsgG =	'**Error - No files to process'				+	`
					"`r`n  No script names found in database."
		Write-Both_ -InputText $strMsgG -Red
		New-ExitProcess_
	}#End If - none
}#End If - No script argument, using database.
Else
{	Write-Verbose_ '  Getting script names from argument...'
	$colFileNamesG = Get-List_ -List $Script:Scripts
}#End Else - using argument list

$intCountSuccessG      = 0
$intCountTotalScriptsG = $colFileNamesG.Count

$strMsgG =	'  Located '			+	`
			$intCountTotalScriptsG	+	`
			' scripts to process.'
Write-Both_ -InputText $strMsgG


ForEach( $strFileNameG IN $colFileNamesG )
{	$strQueryG =	"SELECT * FROM Scripts WHERE FileName = '"	+	`
					$strFileNameG								+	`
					"'"
	$oDataSetG = New-SqlDataset_ -Query $strQueryG
		
	# Iterate through dataset (only one record)
	ForEach( $oRowG IN $oDataSetG.Tables[0] )
	{	$strTempG   = $oRowG.FileName
		$strMsgG    = "    Processing $strTempG"
		Write-Both_ -InputText $strMsgG
		$blnResultG = ProcessRecord	-Record $oRowG
		
		If( $blnResultG ){ $intCountSuccessG ++ }
	}#Next - Row
}#Next - strFileNameG

$strMsgG =	@"
==============================================================================
  Summary
	        Total scripts: $intCountTotalScriptsG
	Successfull processed: $intCountSuccessG
"@
Write-Both_ -InputText $strMsgG

New-ExitProcess_
<#############################################################################
This script was developed to gather scripts from their development locations
  and place them all in a common folder, in a web server file tree.
The script calculates an MD5 hash and writes it into a text file in the
  target location.

The script is designed, primarily, to draw information from a database holding
  metadata about the scripts.  Each time a script is copied, the database is
  updated with a timestamp and the MD5 hash value.
Used without a database, scripts/files can be copied from one common location
  to one target location, with creation of the hash text file for each script.
#############################################################################>
