Param(	#See general comments at end of file.									`
		#List of servers to be processed.										`
		#Individual, comma-separated or file name.								`
		[string]$Servers = ''													`
		#Display string for banner												`
		[string]$VersionInfo = 'Script Name 1.0, September 2018'			,	`
		#Version number															`
		[string]$Version = '2.2.20180907'									,	`
		#Root for log file name, with possible date/time addition				`
		[string]$LogBase = 'Log'											,	`
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
		[string]$AlertRecipient    = 'Trey.Shaffer@ITArtisan.com'			,	`
		[string]$AlertRecipientBCC = ''										,	`
		[string]$AlertRecipientCC  = ''										,	`
		[string]$AlertFrom         = 'PowerShell@ITArtisan.com'				,	`
		[string]$AlertSubject      = '**Alert - '							,	`
		#Switch to cast email in HTML format.									`
		[switch]$AlertHTML													,	`
		#Switch disables sending of email alerts.								`
		[switch]$NoEmail													,	`
		#																		`
		#Secondary/common parameters											`
		#File name only of ps1 file to include									`
		[string]$ModuleFile = 'PsLibrary2'									,	`
		#Path to library module, for use if not available locally.				`
		#  This is the parent folder to the actual module folder.				`
		#  The path to the module will be,										`
		#    <$ModulePath>\<$ModuleFile>\<$ModuleFile>							`
		#Production																`
			[string]$Modulepath = '\\Trey-10-1\PSModules$'					,	`
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
			#[string]$SqlServer = 'Lab-********'							,	`
		#SQL Server database name												`
		[string]$SqlDatabase = 'SystemAdmin'								,	`
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
#File name only of Access database											`
#[string]$MdbFile = 'SomeDatabase.mdb'									,	`
#Path to folder holding Access database										`
#[string]$MdbFolder = ''												,	`
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
	#At this point module is not loaded or avaialable locally.

	If( $ModulePath -EQ '' )
	{	$strMsgL =	'**Error - specified module not '	+	`
					'loaded OR available locally.'		+	`
					"`r`n  Alternate installation "		+	`
					'path not specified.'
		Write-Host $strMsgL -Red
		Exit 77
	}#End If - no ModulePath
	Else
	{	$ModulePath       = EndBackslash $ModulePath
		$ModulePath       = "$ModulePath$ModuleName"
		$strModuleSourceL = "$ModulePath\$ModuleName.psm1"
		
		If( Test-Path $strModuleSourceL )
		{	Write-Verbose '  Alternate source module path confirmed.'
			
			#Module available, load it.
			$strEaHoldL            = $ErrorActionPreference
			$ErrorActionPreference = 'SilentlyContinue'
			Write-Verbose '  Loading module...'
			Import-Module -FullyQualifiedName $strModuleSourceL

			If( $? )
			{	$strMsgL =	'  Required module, '				+	`
							$ModuleName							+	`
							', loaded from alternate location.'
				Write-Both_ -InputText $strMsgL -Red
			}#End If - success
			Else
			{	$strMsgL =	"**Error - Failed to load module, $ModuleName"
				Write-Both_ -InputText $strMsgL -Red
				If( $NoAbort )
				{	$ErrorActionPreference = $strEaHoldL
					Return $False
				}#End If
				Else{ Exit 77 }
			}#End If - error
			$ErrorActionPreference = $strEaHoldL
			Return $True
		}#End If - Alternate module found
		Else
		{	$strMsgL =	'**Error - Alternate module source not found.'	+	`
						"`r`n`t$strModuleSourceL"
			Write-Host $strMsgL -ForegroundColor 'Red'
			Exit 77
		}#End Else - not found
	}#End Else - Module path supplied
}# End Local Load Module #####################################################



##############################################################################
###Main
######################################################################
If( $True ){ #Common setup code
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

##Enable this section if using any database access.
#$Global:oDbConnectionG = New-SqlDbConnection_	-Database $Script:SqlDatabase	`
#												-Instance $Script:SqlInstance	`
#												-Port     $Script:SqlPort		`
#												-Pwd      $Script:SqlPwd		`
#												-Server   $Script:SqlServer		`
#												-UserID   $Script:SqlUser

##Enable this if using email.
##Requires database access to configure email defaults from database.
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

##Enable for Exchange 2007
##Can't load SnapIn from within a module.
#$strSnapInG = 'Microsoft.Exchange.Management.PowerShell.Admin'
#If( ( Get-PSSnapin	-Name $strSnapInG -EA 'SilentlyContinue' ) -EQ $null )
#{	Add-PsSnapin	-Name $strSnapInG -EA 'SilentlyContinue' }

}#End common base code
##############################################################################
#Code for this script begins here...
If( $Script:Servers -eq '' )
{ $colServersG = @( $ENV:ComputerName ) }
Else
{	$colServersG = Get-List_ -List $Script:Servers }


##############################################################################
#Typical prep for email notification...
#If( $Script:NoEmail -EQ $False )
#{	
#	$Nothing = Invoke-SendEmail_									`
#							-BccTo       $Script:AlertRecipientBCC	`
#							-CcTo        $Script:AlertRecipientCC	`
#							-MsgBody     $strMsgBodyG				`
#							-UserFrom    $Script:AlertFrom			`
#							-Recipient   $Script:AlertRecipient		`
#							-Subject     $Script:AlertSubject		`
#							-Attachments $strAttachmentG
#}#End If - not NoEmail
##############################################################################


New-ExitProcess_
<#############################################################################
Block comments

V 2.2.20180907 - Simplified LoadModuleLocal

#############################################################################>
