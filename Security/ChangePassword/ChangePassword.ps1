Param(	#See general comments at end of file.									`
		#List of servers to be processed.										`
		#Individual, comma-separated or file name.								`
		[string]$Servers = 'D:\PS\ServerLists\PasswordChange.txt'			,	`
		#List of user acount names to be modified.								`
		[string]$Users = 'Administrator'									,	`
		#Method to generate password.											`
		#  Random - Uses .Net password generator.								`
		#  Simple - Concatenates -Seed with computername.						`
		[switch]$Random														,	`
		[switch]$Simple														,	`
		#Uniform/Unique determines whether same password is used for same		`
		#  username on each computer.											`
		[switch]$Uniform													,	`
		[switch]$Unique														,	`
		#String value to concatenate with attributes of target computer(s).		`
		[string]$Seed = 'Shaffer1'											,	`
		#Total length of generated random password component.					`
		[int]$PwdLengthOverall = 12											,	`
		#Number of non-alpha characters in generated password.					`
		[int]$PwdCountNonAlpha = 4											,	`
		#																		`
		#Display string for banner												`
		[string]$VersionInfo = 'Change Password 2.0, October 2018'			,	`
		#Root for log file name, with possible date/time addition				`
		[string]$LogBase = 'Pwd'											,	`
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
		#Path to library module, for use if not available locally				`
		#Production																`
			[string]$Modulepath = '\\Trey-10-1\PSModules$'					,	`
		#Lab																	`
			#\\LabMgr01\PSModules$'											,	`
		#Minimum acceptable version of library File								`
		[string]$ModuleMinVersion = '1.0'									,	`
		#Default delimiter to use in lists										`
		[string]$Delimiter = ','											,	`
		#Host name of SQL Server												`
		#Production																`
			[string]$SqlServer = 'Sql17-01'									,	`
		#Lab																	`
			#[string]$SqlServer = 'Lab-Sql01'								,	`
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
		)##############################################################################
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
Creates password string.

.Description
Simple concatenation of two inputs, computer name and seed string.
#############################################################################>
Function New-Password()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.String" )]
Param(	[string]$User							,	`
		[hashTable]$UniformHT					,	`
		[string]$Seed         = $Script:Seed	,	`
		[switch]$Simple       = $Script:Simple	,	`
		[switch]$Uniform      = $Script:Uniform		`
		)
	If( $Simple ){ $strPwdL = $Seed }
	Else
	{	If( $Uniform ){ $strPwdL = $UniformHT.$User }
		Else { $strPwdL = New-RandomPassword }
	}#End Else - not simple
	Return $strPwdL
}#End -  New Password ########################################################



<#############################################################################
.Synopsis
Creates password using .Net

.Description
Create random password of specified length/complexity.
#############################################################################>
Function New-RandomPassword()
{#############################################################################
[CmdletBinding()]
[OutputType( "System.String" )]
Param(	[int]$LengthOverall  = $Script:PwdLengthOverall	,	`
		[int]$LengthNonAlpha = $Script:PwdCountNonAlpha		`
		)
	$strPwdL =	[system.web.security.membership]::	`
				GeneratePassword( $LengthOverall, $LengthNonAlpha )
	Return $strPwdL
}#End - New Random password ##################################################



##############################################################################
###Main
######################################################################
If( $True ){ #Common setup code
CLS
Trap { New-ErrorInfo_ -Error $_ }

Set-PsDebug -Strict
$Global:dteStartG             = Get-date
$Global:blnLogInitializedG    = $False
$Script:ErrorActionPreference = 'SilentlyContinue'
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

##Enable this if using email.  Requires database access to configure email.
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
#Validate arguments
#Random overrides Simple.
If( $Script:Random ){ $Script:Simple = $False }

#Random as default if not specified.
If(	$Script:Random -EQ $False -AND $Script:Simple -EQ $False )
{	$Script:Random = $True }

#Unique overrides Uniform.
If( $Script:Unique ){ $Script:Uniform = $False }

#Unique as default if not specified.
If(	$Script:Unique -EQ $False -AND $Script:Uniform -EQ $False )
{	$Script:Uniform = $True }

$strMsgG = @"
  Argument values:
     Simple: $Script:Simple
     Random: $Script:Random
    Uniform: $Script:Uniform
     Unique: $Script:Unique
"@
Write-Both_ $strMsgG

If( $Script:Servers -eq '' )
{ $colServersG = @( $ENV:ComputerName ) }
Else
{	$colServersG = Get-List_ -List $Script:Servers }

$colUsersG = Get-List_ -List $Script:Users

If( $Script:Uniform )
{	#Create hashtable of user names and passwords.
	$htPasswordsG = @{}
	
	If( $Script:Random )
	{	ForEach( $strUserG IN $colUsersG )
		{	$strPasswordG = New-RandomPassword
			$htPasswordsG.$strUserg = $strPasswordG
		}#Next - strUserG
	}#End If - random
	Else
	{	ForEach( $strUserG IN $colUsersG )
		{	$htPasswordsG.$strUserg = $Script:Seed }
	}#End Else - not random
}#End If - uniform

$intCountFailedConnectG = 0
ForEach( $strComputerG IN $colServersG )
{	$strMsgG = "  Processing computer, $strComputerG..."
	Write-Both_ -InputText $strMsgG
	
	If( Test-Connection $strComputerG )
	{	ForEach( $strUserG IN $colUsersG )
		{	$strMsgG =	"    User account, $strUserG..."
			Write-Host $strMsgG
			
			$oUserG = $Null
			$oUserG = [adsi]$oUuserG = “WinNT://$Computer/$strUserG”
			
			If( $oUserG )
			{	$strPasswordG = New-Password	-User $strUserG			`
												-UniformHT $htPasswordsG
				
				$strMsgG = "  $strComputerG   $strUserG  $strPasswordG"
				Write-Both_ $strMsgG
				
				$oUserG.SetPassword( $strPasswordG )

				####([adsi]"WinNT://$strComputerG/$strUserG").SetPassword( $strPasswordG )
			}#End If - user found
			Else
			{	$strMsgG =	'**Error - Failed to find user account.'	+	`
							"`r`n  Unable to locate user account, "		+	`
							"$strUserG on computer, $strComputerG."
				Write-Both_ $strMsgG -Red
			}#End Else - user not found
		}#Next - User
	}#End If - Test-Connection
	Else
	{	$intCountFailedConnectG ++
		$strMsgG =	'**Error - Failed to connect to target computer.'	+	`
					"'  Unable to connect to, $strComputerG."
		Write-Both_ $strMsgG -Red
	}#End Else - no connection
}#Next - strComputerG

New-ExitProcess_
<#############################################################################
Block comments


#############################################################################>
