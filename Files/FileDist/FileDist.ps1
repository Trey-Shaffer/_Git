Param(	#See general comments at end of file.									`
		#Configuration file.  File name only.  Assumed to be in ConfigPath.		`
		#This is only arg needed if config file is complete.					`
		[string]$ConfigFile = 'PopulateCustomAttributes.xml'				,	`
		#Default path for config file.											`
		[string]$ConfigPath =	'C:\Scripts\Files\FileDist\Configuration\'	,	`
		#																		`
		#Input for server list.													`
		$ServerList = $Null													,	`
		#Script default server list.											`
		[string]$strServerListG = ''										,	`
		#Input for computer holding source file(s).								`
		$SourceComputer = $Null												,	`
		#Script default for source computer.									`
		[string]$strSourceComputerG = ''									,	`
		#Input for name of share holding source file(s).						`
		$SourceShare = $Null												,	`
		#Script default for source share.										`
		[string]$strSourceShareG = ''										,	`
		#Input for path, below source share, to source file(s).					`
		$SourcePath = $Null													,	`
		#$Script default for source path.										`
		[string]$strSourcePathG = ''										,	`
		#Input for share name on all target computers.							`
		$TargetShare = $Null												,	`
		#Script default for target share.										`
		[string]$strTargetShareG = 'Staging$'								,	`
		#Input for target path, below TargetShare.								`
		$TargetPath = $Null													,	`
		#Script default for target path.										`
		[string]$strTargetPathG = ''										,	`
		#Deleting target content refers to removing the entire contents of		`
		#  the target folder, not just overwriting files.						`
		#Enabling deletes content AND copies new material.						`
		[switch]$DeleteContent												,	`
		#Script default value for delete target content.						`
		[bool]$blnDeleteContentG = $False									,	`
		#DeleteOnly removes content and exits without copying any new files.	`
		#Again, this delets ALL content in the target folder.					`
		[switch]$DeleteContentOnly											,	`
		#Script default for delete content only.								`
		[bool]$blnDeleteContentOnlyG = $False								,	`
		#Delete content of target folder and all sub-folders.					`
		[switch]$Recurse													,	`
		#Script default, delete content of target folder and all sub-folders.	`
		[bool]$blnRecurseG													,	`
		#Input for Create (folders in target path)								`
		[switch]$Create														,	`
		#Script default for Create.												`
		[bool]$blnCreateG = $False											,	`
		#Pre/PostCopyCommand not intended as command line args.					`
		[string]$strPreCopyCommandG  = ''									,	`
		[string]$strPostCopyCommandG = ''									,	`
		#																		`
		#Display string for banner												`
		[string]$VersionInfo = 'File Distribution 2.7, April 2017'			,	`
		#Root for log file name, with possible date/time addition				`
		[string]$LogBase = 'FileDist'										,	`
		#Path to folder holding script log file.  '' = script folder			`
		[string]$LogFolder = ''												,	`
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
			[string]$SqlServer = 'SqlCluster01'								,	`
		#Lab																	`
			#[string]$SqlServer = 'Lab-SqlClus01'							,	`
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
Function CopyFiles()
{#############################################################################
Param(	[string]$SourceFqPath			,	`
		[string]$TargetComputer			,	`
		[string]$TargetShare			,	`
		[string]$TargetPath				,	`
		$PreCopyCommand         = $Null	,	`
		$PostCopyCommand        = $Null		`
		)
	#SourceFqPath not validated here.  Assume already tested...

	$blnReadyToCopyL        = $False
	$strTargetShareL        = '\\' + $TargetComputer + '\' + $TargetShare
	$strTargetShareAndPathL = $TargetShare	#Used for Pre/PostCopyCommand

	If( Test-Path $strTargetShareL )
	{	Write-Verbose_ -InputText '  Target share exists.'

		If( $TargetPath )
		{	#Path specified beyond share.
			#Add path to share UNC.
			Write-Verbose_ -InputText "  Target path specified in addition to share."
			$TargetPath              = $TargetPath.TrimStart( '\' )
			$strFullTargetPathL      = "$strTargetShareL\$TargetPath"
			$strTargetShareAndPathL += "\$TargetPath"	#For Pre/PostCopyCommand
		}#End If - TargetPath
		Else
		{	$strFullTargetPathL = $strTargetShareL }#End Else - not target path

		If( Test-Path $strFullTargetPathL ){ $blnReadyToCopyL = $True }
		Else
		{	If( $Script:blnCreateG )
			{	$strMsgL =	'**Warning - path not found.'	+	`
							"`r`n`t$strFullTargetPathL"
				Write-Both_ -InputText     $strMsgL -Color 'Yellow|Black'
				Write-Verbose_ -InputText '  Creating target path...'
				$blnResultL = CreatePath -Path $strFullTargetPathL

				If( $blnResultL )
				{	Write-Verbose_ '    Created.'
					$blnReadyToCopyL = $True
				}#End If - blnResultL
				Else
				{	$strMsgL =	'**Error - unable to complete target path.'
					Write-Both_ -InputText $strMsgL -Red
				}#End Else - unable to create full target path
			}#End If - Create
			Else
			{	$strMsgL =	'**Error - path not found.'		+	`
							"`r`n`t"						+	`
							$strFullTargetPathL				+	`
							"`r`n"							+	`
							'  Create option not specified.'
				Write-Both_ -InputText  $strMsgL -Red
				Return $False
			}#End Else
		}#End If - path not found
	}#End If - Target share exists
	Else
	{	$strMsgL =	'**Error - Target share not found.'		+	`
					"`r`n  Share, $TargetShare, not found."	+	`
					"`r`n  Skipping this host..."
		Write-Both_ -InputText $strMsgL -Red
	}#End Else - target share not found

	If( $blnReadyToCopyL )
	{	If( $PreCopyCommand -ne '' )
		{	$PreCopyCommand =													`
						$PreCopyCommand.Replace( '*Computer*', $TargetComputer )

			$PreCopyCommand =														`
				$PreCopyCommand.Replace( '*TargetFolder*', $strTargetShareAndPathL )

			Write-Verbose_ -InputText											`
								"  Executing pre-copy command, $PreCopyCommand"

			$oScriptBlockL = [ScriptBlock]::Create( $PreCopyCommand )
			Invoke-Command	-ComputerName $TargetComputer	`
							-ScriptBlock  $oScriptBlockL
		}#End If - PreCopyCommand

		$strMsgL =  '    Copying files to: ' + $strFullTargetPathL
		Write-Both_ -InputText  $strMsgL
		Copy-Item	-Path        "$SourceFqPath\*"		`
					-Destination $strFullTargetPathL	`
					-EA          'Continue'				`
					-Recurse							`
					-Force

		If( $PostCopyCommand -ne '' )
		{	$PostCopyCommand =													`
						$PostCopyCommand.Replace( '*Computer*', $TargetComputer )

			$PostCopyCommand =														`
				$PostCopyCommand.Replace( '*TargetFolder*', $strTargetShareAndPathL )

			Write-Verbose_ -InputText											`
								"  Executing post copy command, $PostCopyCommand"

			$oScriptBlockL = [ScriptBlock]::Create( $PostCopyCommand )
			Invoke-Command	-ComputerName $TargetComputer	`
							-ScriptBlock  $oScriptBlockL
		}#End If - PostCopyCommand

		Return $True
	}#End If - blnReadyToCopyL
	Else{ Return $False }
}#End Copy Files #############################################################


##############################################################################
Function CreatePath()
{#############################################################################
Param(	[string]$Path			,	`
		[switch]$AbortOnFailure		`
		)
	$Path          = $Path.TrimStart( '\' )
	$colFullPathL  = $Path.Split( '\' )
	$strRootL      = $colFullPathL[0]
	$colPathPartsL = @()

	If( $strRootL.Contains( ':' ) )
	{	Write-Verbose_ -InputText "  Using drive letter, $strRootL as root."
		If( $colFullPathL.Length -gt 1 )
		{	For(	$int1L = 1						;	`
					$int1L -le $colFullPathL.Length	;	`
					$int1L ++							`
				)
			{	$colPathPartsL += ,$colFullPathL[ $int1L ] }
		}#End If - gt 1
	}#End If - drive letter
	Else
	{	#Root must be a computer name for UNC format path.
		Write-Verbose_ -InputText "  Using hostname, $strRootL."
		If( Test-Connection( $strRootL ) )
		{	#Replace leading '\\' and add share name.
			$strRootL = "\\$strRootL\" + $colFullPathL[1]
			Write-Verbose_ -InputText "  Using hostname\share, $strRootL as root."

			If( $colFullPathL.Length -gt 2 )
			{	For(	$int1L = 2						;	`
						$int1L -le $colFullPathL.Length	;	`
						$int1L ++							`
					)
				{	$colPathPartsL += ,$colFullPathL[ $int1L ] }
			}#End If - gt 2
		}#End If - connect
		Else
		{	$strMsgL =	"  Failed to connect to $strRootL."	+	`
						'  Verify network connectivity.'
			Write-Both_ -InputText  $strMsgL -Red
			If( $AbortOnFailure )
			{	Write-Both_ -InputText  '  Terminating processing...' -Red
				New-ExitProcess_
			}#End If
			Else{ Return $False }
		}#End Else - connect
	}#End Else - MS-DOS/UNC

	$strPathL = $strRootL

	ForEach( $strItemL IN $colPathPartsL )
	{	$strTargetL = "$strPathL\$strItemL"

		If( Test-Path( $strTargetL ) )
		{	Write-Verbose "  Path found: $strTargetL" }
		Else
		{	Write-Verbose "  Creating: $strTargetL..."
			New-Item	-Path     $strPathL			`
						-Name     $strItemL			`
						-ItemType 'Directory'		`
						-EA       'SilentlyContinue'
		}#End Else - no connect
		$strPathL = $strTargetL
	}#Next - strItemL
	Return $True
}#End - Create Path ##########################################################


##############################################################################
Function DeleteContent()
{#############################################################################
Param(	[array]$ServerList          = $Script:colTargetServersG			,	`
		[string]$TargetShareAndPath = $Script:strTargetShareAndPathG	,	`
		[bool]$Recurse              = $Script:Recurse						`
		)
	$TargetShareAndPath = $TargetShareAndPath.TrimStart( '\' )

	ForEach( $strServerL IN $ServerList )
	{	$strTargetL =	'\\'		+	`
				$strServerL			+	`
				'\'					+	`
				$TargetShareAndPath
		$strMsgL =	"  Deleting contents of $strTargetL..."
		Write-Both_ -InputText  $strMsgL

		$strTargetL += '\*'

		If( $Recurse )
		{
		$strMsgL =	"    Delete includes sub-folders and files."
		Write-Both_ -InputText  $strMsgL
		Remove-Item	-Path    $strTargetL	`
					-Recurse				`
					-Confirm:$False			`
					-Force					`
					-EA      'Continue'
		}#End If - Recurse
		Else
		{	Remove-Item	-Path    $strTargetL	`
						-Confirm:$False			`
						-Force					`
						-EA      'Continue'
		}#End Else
	}#Next - strServerG
}#End - Delete Content #######################################################


##############################################################################
Function DisplayScriptDefaults( )
{#############################################################################
	$strMsgL =	"Script default values..."								+	`
				"`r`n`tInput file version: $Script:InputFileMinVersionG"
	Write-Verbose_ -InputText $strMsgL

	#If( $Script:strConfigFileG -eq '' )
	If( $Script:ConfigFile -eq '' )
	{	$strTempL = 'None'
	}#End If
	Else{ $strTempL = $Script:ConfigFile }
	$strMsgL =	" Configuration file: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:strServerListG -eq '' )
	{	$strTempL = 'None'
	}#End If
	Else{ $strTempL = $Script:strServerListG }
	$strMsgL =	"        Server list: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:strSourceComputerG -ne '' )
	{	$strTempL = $Script:strSourceComputerG }
	Else{ $strTempL = 'None' }
	$strMsgL =	"    Source computer: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:strSourceShareG -ne '' )
	{	$strTempL = $Script:strSourceShareG }
	Else{ $strTempL = 'None' }
	$strMsgL =	"       Source share: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:strSourcePathG -ne '' )
	{	$strTempL = $Script:strSourcePathG }
	Else{ $strTempL = 'None' }
	$strMsgL =	"        Source path: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:strTargetShareG -ne '' )
	{	$strTempL = $Script:strTargetShareG }
	Else{ $strTempL = 'None' }
	$strMsgL =	"       Target share: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:strTargetPathG -ne '' )
	{	$strTempL = $Script:strTargetPathG }
	Else{ $strTempL = 'None' }
	$strMsgL =	"        Target path: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:blnDeleteContentG )
	{	$strTempL = 'True' }
	Else{ $strTempL = 'None' }
	$strMsgL =	"     Delete content: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:blnDeleteContentOnlyG )
	{	$strTempL = 'True' }
	Else{ $strTempL = 'None' }
	$strMsgL =	"Delete content only: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:blnRecurseG )
	{	$strTempL = 'True' }
	Else{ $strTempL = 'None' }
	$strMsgL =	"Delete recursively: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:blnCreateG )
	{	$strTempL = 'True' }
	Else{ $strTempL = 'None' }
	$strMsgL =	" Create target path: $strTempL"
	Write-Verbose_ -InputText $strMsgL
}#End Display Script Defaults ################################################


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
Apply command line argument to script configuration variables.

.Description
If a given command line argument exists, assign its valu to relevant variable.
#############################################################################>
Function ProcessCommandLine()
{#############################################################################
	$strMsgL =	"`r`n`r`n    Processing command line arguments..."
	Write-Verbose_ -InputText $strMsgL

	If( $Script:ConfigFile )
	{	$Script:strConfigFileG = [string]$Script:ConfigFile
		$strTempL = $Script:strConfigFileG
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = " Configuration file: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:ServerList )
	{	$Script:strServerListG = [string]$Script:ServerList
		$strTempL = $Script:strServerListG
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = "        Server list: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If(	$Script:SourceComputer )
	{	$Script:strSourceComputerG = [string]$Script:SourceComputer
		$strTempL = $SCript:strSourceComputerG
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = "    Source computer: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If(	$Script:SourceShare )
	{	$Script:strSourceShareG = [string]$Script:SourceShare
		$strTempL = $Script:strSourceShareG
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = "       Source share: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If(	$Script:SourcePath )
	{	$Script:strSourcePathG = [string]$Script:SourcePath
		$strTempL = $Script:strSourcePathG
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = "        Source path: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If(	$Script:TargetShare )
	{	$Script:strTargetShareG = [string]$Script:TargetShare
		$strTempL = $Script:strTargetShareG
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = "       Target share: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If(	$Script:TargetPath )
	{	$Script:strTargetPathG = [string]$Script:TargetPath
		$strTempL = $Script:strTargetPathG
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = "        Target path: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:DeleteContent )
	{	$Script:blnDeleteContentG = $True
		$strTempL = 'True'
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = "     Delete content: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:DeleteContentOnly )
	{	$Script:blnDeleteContentOnlyG = $True
		$strTempL = 'True'
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = "Delete content only: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:Recurse )
	{	$Script:blnRecurseG = $True
		$strTempL = 'True'
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = " Delete recursively: $strTempL"
	Write-Verbose_ -InputText $strMsgL

	If( $Script:Create )
	{	$Script:blnCreateG = $True
		$strTempL = 'True'
	}#End If
	Else{ $strTempL = 'None' }
	$strMsgL = " Create target path: $strTempL"
	Write-Verbose_ -InputText $strMsgL
}#End - Process Command Line #################################################


<#############################################################################
.Synopsis
Configure/override script variables based on coresponding values from input file.

.Description
Get all active values from input file.
For each potential value, if it exists from input file, assign that value
  to the coresponding Script-scope variable.

#############################################################################>
Function ProcessConfigFile()
{#############################################################################
Param(	[string]$ConfigurationFile )

	If( ( Test-Path $ConfigurationFile ) -EQ $False )
	{	$strMsgL =	'**Error - File not found.'			+	`
					"`r`n  Configuration file, "		+	`
					$ConfigurationFile					+	`
					', not found.'
		Write-Both_ -InputText $strMsgL -Red
		New-ExitProcess_
	}#End If - file not found
	
	#Read input file and convert it to XML document.
	[xml]$xmlFileL = Get-Content $ConfigurationFile

	$strTempG = $xmlFileL.Root.Verbose
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:Verbose2 = $True
		$Script:VerbosePreference = 'Continue'
		Write-Verbose_ -InputText "            Verbose: True"
	}#End If

	$strMsgL =	"`r`n`r`n  Input values from file..."	+	`
				"`r`n      $ConfigurationFile..."
	Write-Verbose_ -InputText $strMsgL

	$strTempG = $xmlFileL.Root.FileVersion
	$strTempG = [string]$strTempG
	If( ( $strTempG -lt $Script:InputFileMinVersionG ) )
	{	$strMsgL =	"  The version of the specified config file:`r`n`t"	+	`
					$Script:ConfigFile									+	`
					"`r`n    is "										+	`
					$strTempG											+	`
					', less than the required version, '				+	`
					$Script:InputFileMinVersionG						+	`
					'.'
		Write-Both_ -InputText  $strMsgL -Red
		New-ExitProcess_ 99
	}#End - not min version
	$strMsgL =	'            Version: '	+	`
				$strTempG				+	`
				' (OK)'
	Write-Verbose_ -InputText $strMsgL

	$strTempG = $xmlFileL.Root.ServerList
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:strServerListG = $strTempG
		$strMsgL =	"        Server list: $Script:strServerListG"
		Write-Verbose_ -InputText $strMsgL
	}#End If

	$strTempG = $xmlFileL.Root.SourceComputer
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:strSourceComputerG = $xmlFileL.Root.SourceComputer.Trim()
		$strMsgL =	"    Source computer: $Script:strSourceComputerG"
		Write-Verbose_ -InputText $strMsgL
	}#End If

	$strTempG = $xmlFileL.Root.SourceShare
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:strSourceShareG = $xmlFileL.Root.SourceShare.Trim()
	$strMsgL =	"       Source share: $Script:strSourceShareG"
	Write-Verbose_ -InputText $strMsgL
	}#End If

	$strTempG = $xmlFileL.Root.SourcePath
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:strSourcePathG = $xmlFileL.Root.SourcePath.Trim()
		$strMsgL =	"        Source path: $Script:strSourcePathG"
		Write-Verbose_ -InputText $strMsgL
	}#End If

	$strTempG = $xmlFileL.Root.TargetShare
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:strTargetShareG = $xmlFileL.Root.TargetShare.Trim()
		$strMsgL =	"       Target share: $Script:strTargetShareG"
		Write-Verbose_ -InputText $strMsgL
	}#End If

	$strTempG = $xmlFileL.Root.TargetPath
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:strTargetPathG = $xmlFileL.Root.TargetPath.Trim()
		$strMsgL =	"        Target path: $Script:strTargetPathG"
		Write-Verbose_ -InputText $strMsgL
	}#End If

	$strTempG = $xmlFileL.Root.DeleteContent
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:blnDeleteContentG = $True
		Write-Verbose_ -InputText "     Delete content: True"
	}#End If

	$strTempG = $xmlFileL.Root.DeleteContentOnly
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:blnDeleteContentOnlyG = $True
		Write-Verbose_ -InputText "Delete content only: True"
	}#End If

	$strTempG = $xmlFileL.Root.Recurse
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:blnRecurseG = $True
		Write-Verbose_ -InputText " Delete recursively: True"
	}#End If

	$strTempG = $xmlFileL.Root.Create
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:blnCreateG = $True
		Write-Verbose_ -InputText " Create target path: True"
	}#End If

	$strTempG = $xmlFileL.Root.PreCopyCommand
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:strPreCopyCommandG = $xmlFileL.Root.PreCopyCommand.Trim()
		$strMsgL =	"   Pre-copy command: $Script:strPreCopyCommandG"
		Write-Verbose_ -InputText $strMsgL
	}#End If

	$strTempG = $xmlFileL.Root.PostCopyCommand
	$strTempG = [string]$strTempG
	If( $strTempG.Trim() -NE '' )
	{	$Script:strPostCopyCommandG = $xmlFileL.Root.PostCopyCommand.Trim()
		$strMsgL =	"  Post-copy command: $Script:strPostCopyCommandG"
		Write-Verbose_ -InputText $strMsgL
	}#End If
}#End Process Config File ####################################################



##############################################################################
###Main
######################################################################
If( $True ){ #Common setup code
CLS
Trap { New-ErrorInfo_ -Error $_ }

Set-PsDebug -Strict
$Global:dteStartG             = Get-date
$Script:ErrorActionPreference = 'Continue'
$Script:ErrorDescriptionG     = 'No additional information'
$Script:arrLogHoldG           = @( "First Line" )
$Script:blnLogInitializedG    = $False
$Script:intReturnCodeG        = 0
$Script:strScriptPathG        = GetScriptPath
$Script:strSeparatorG         = ( '=' * 78 ) + "`r`n"
$Script:strThisScriptG        = Split-Path -Path $Script:strScriptPathG -Leaf

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
$Script:strLogFileFqG = Initialize-Log_	-LogAppend $Script:LogAppend	`
										-LogBase   $Script:LogBase		`
										-LogExt    $Script:LogExt		`
										-LogDate   $Script:LogDate		`
										-LogFolder $Script:LogFolder	`
										-LogTime   $Script:LogTime
If( $Script:strLogFileFqG ) { $Script:blnLogInitializedG = $True }

##Enable this section if using any database access.
#$Global:oDbConnectionG = New-SqlDbConnection_	-Database $Script:SqlDatabase	`
#												-Instance $Script:SqlInstance	`
#												-Port     $Script:SqlPort		`
#												-Pwd      $Script:SqlPwd		`
#												-Server   $Script:SqlServer		`
#												-UserID   $Script:SqlUser

##Enable this if using email.  Requires database access to configure email.
#$Script:oSmtpServerG = Get-SmtpServer_ $Script:oDbConnectionG

Write-Banner_ -VersionInfo $VersionInfo  -Separator $strSeparatorG
}#End common base code
##############################################################################
$Script:InputFileMinVersionG = '2.0'
DisplayScriptDefaults

If( $Script:ConfigFile.Contains( '\' ) )
{	#Assume this is fully qualified file name.  Use as is.
	ProcessConfigFile $Script:ConfigFile
}#End If
Else
{	#Config file name only.
	$Script:strConfigFileG  = EndBackslash $SCript:ConfigPath
	$Script:strConfigFileG += $Script:ConfigFile
	ProcessConfigFile       $Script:strConfigFileG
}#End Else

ProcessCommandLine

$strMsgG = @"

  Configuration Values Used
            Server list: $Script:strServerListG
     Configuration file: $Script:strConfigFileG
        Source computer: $Script:strSourceComputerG
           Source share: $Script:strSourceShareG
            Source path: $Script:strSourcePathG
           Target share: $Script:strTargetShareG
            Target path: $Script:strTargetPathG
         Delete content: $Script:blnDeleteContentG
    Delete content only: $Script:blnDeleteContentOnlyG
     Create target path: $Script:blnCreateG

"@
Write-Verbose_ -InputText $strMsgG

$colTargetServersG = Get-ServerList_ -List $Script:strServerListG  -NetBIOS
$colTargetServersG = $colTargetServersG | Sort-Object

#Remove any backslashes included with share names.
$Script:strTargetShareG = $Script:strTargetShareG.Replace( "\", '' )
$Script:strSourceShareG = $Script:strSourceShareG.Replace( "\", '' )

If( $Script:strTargetPathG -ne '' )
{	$Script:strTargetShareAndPathG = "$Script:strTargetShareG\$Script:strTargetPathG" }
Else{	$Script:strTargetShareAndPathG = $Script:strTargetShareG }

If( $Script:blnDeleteContentOnlyG	`
	-OR									`
	$Script:blnDeleteContentG		`
	)
{	$strMsgG =	'  Warning - All content of the target folder, '	+	`
				'on all target machines, will be deleted.'
	If( $Script:Recurse )
	{	$strMsgG +=	"`r`n    Delete will include sub-folders." }
	$strMsgG += "`r`n  Enter (Y)es to continue."

	$strResponseG = $Null
	While( $strResponseG -eq $Null )
	{	$strResponseG = Read-Host $strMsgG
		If( $strResponseG -ne $Null )
		{	If( $strResponseG.ToUpper() -Contains 'Y' ){}	#Do nothing
			Else
			{	#Response not (Y)es.
				New-ExitProcess_ 88
			}#End Else - No
		}#End If - not null
	}#WEnd If - Response not null
}#End If - Delete

If( $Script:blnDeleteContentOnlyG )
{	DeleteContent	-ServerList $Script:colTargetServersG		`
					-TargetPath $Script:strTargetShareAndPathG	`
					-Recurse    -Script:blnRecurseG
	New-ExitProcess_
}#End If - delete only

If( $Script:strSourceComputerG -eq $Null )
{	$strMsgG =	'  No Source Computer name supplied.  Using local host.'
	Write-Verbose_ -InputText $strMsgG

	$oIpG = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()
	$strFqHostG = $oIpG.HostName + '.' + $oIpG.DomainName
}#End If - no source computer name
Else {$strFqHostG = $Script:strSourceComputerG.Replace( '\', '' ) }
$strFullSourcePathG = '\\' + $strFqHostG + '\' + $Script:strSourceShareG

$strMsgG =	" `r`n"					+	`
			'  Source computer: '	+	`
			$strFqHostG				+	`
			"`r`n"					+	`
			'     Source share: '	+	`
			$Script:strSourceShareG

If( $Script:strSourcePathG -ne '' )
{	$strMsgG +=	"`r`n"					+	`
				'      Source path: '	+	`
				$Script:strSourcePathG
}#End If

$strMsgG +=	"`r`n"					+	`
			"`r`n"					+	`
			'     Target share: '	+	`
			$Script:strTargetShareG

If( $Script:strTargetPathG -ne '' )
{	$strMsgG +=	"`r`n      Target path: "	+	`
			$Script:strTargetPathG			+	`
			"`r`n`r`n     Source share: "	+	`
			$strFullSourcePathG				+	`
			"`r`n`r`n      Server list: "	+	`
			$Script:strServerListG			+	`
			"`r`n`r`n"
	Write-Both_ -InputText  $strMsgG
}#End If - not empty

#Check the source \\computer\share
If( ( Test-Path $strFullSourcePathG ) -eq $False )
{	Write-Both_ -InputText  "**Error - Source share not found, \$SourceShare" -Red
	New-ExitProcess_ 77
}#End If

If( $Script:strSourcePathG )
{	#Have path in addition to source share.
	Write-Verbose_ -InputText '  Source path specified.'
	If( $Script:strSourcePathG.StartsWith( '\' ) )
	{	#Remove leading \
		$Script:strSourcePathG = $Script:strSourcePathG.TrimStart( '\' )
	}#End If - starts with \

	#Make array of folder names in path.
	$arrPathG = $Script:strSourcePathG.Split( '\' )

	#Confirm path one folder at a time.
	ForEach( $strPathG IN $arrPathG )
	{	$strFullSourcePathG += "\$strPathG"
		If( -NOT( Test-Path $strFullSourcePathG ) )
		{	$strMsgG =	'**Error - Source path not found.'	+	`
						"`r`n`t"							+	`
						$strFullSourcePathG
			Write-Both_ -InputText  $strMsgG -Red
		}#End If - path not found
	}#Next - path folder
	$strMsgG = '  Full path to source: ' + $strFullSourcePathG
	Write-Verbose_ -InputText $strMsgG
}#End If - Source path exists

$strMsgG =	'  This process will copy all files and folders from: '	+	`
			"`r`n`t"												+	`
			$strFullSourcePathG										+	`
			"`r`n`r`n"												+	`
			'  To '													+	`
			$Script:strTargetShareG
If( $Script:strTargetPathG ){ $strMsgG += "\$Script:strTargetPathG" }
$strMsgG += ' on the following computers:'

ForEach( $strServerG IN $colTargetServersG )
{	$strMsgG += "`r`n`t" + $strServerG }

$strMsgG +=	"`r`n`r`n"											+	`
			'  Existing files, of same name, on target will '	+	`
			'be overwritten.'									+	`
			"`r`n`r`n"
If( $Script:blnDeleteContentG )
{	$strMsgG +=	"  Target folder content will be deleted prior to copy.`r`n`r`n"
	If( $Script:Recurse )
	{	$strMsgG +=	"    Delete will include sub-folders.`r`n`r`n" }
}#End If
Write-Both_ -InputText  $strMsgG

ForEach( $strServerG IN $colTargetServersG )
{	Write-Both_ -InputText "`r`n  Processing $strServerG..."
	If( $Script:blnDeleteContentG )
	{	#Create array of one Server
		$arrServersG = @( $strServerG )

		DeleteContent	-ServerList $arrServersG							`
						-TargetShareAndPath $Script:strTargetShareAndPathG	`
						-Recurse            $Script:Recurse
	}#End If - DeleteContent

	$Nothing = CopyFiles	-SourceFqPath    $strFullSourcePathG			`
							-TargetComputer  $strServerG					`
							-TargetShare     $Script:strTargetShareG		`
							-TargetPath      $Script:strTargetPathG			`
							-PreCopyCommand  $Script:strPreCopyCommandG		`
							-PostCopyCommand $Script:strPostCopyCommandG	`
							-EA              'Continue'
}#Next - strServerG

New-ExitProcess_
<#############################################################################
File copy/distribution utility.

Parameters:
	-SourceComputer: Defaults to local.  Must be fully qualified.
	-SourceShare: Must reference a share name on SourceComputer.
	-SourcePath: Optional, path beyond SourceShare.
	-TargetShare: The name of a share expected on all target hosts.
	-TargetPath: Optional, path beyond TargetShare.
	-Input file:  UNC path to text file holding list of servers/targets
	  for copy actions.  File in same folder as script needs name only.

Input file format - XML.

Stage files for distribution in the SourceFolder specified.
All files and folders in SourceFolder will be copied to targets.
Existing files, of same name, on target will be overwritten.

Command Line:  Parameters may be specified on the command line.
               Default values are coded at begining of script.

V 1.6 added create path

V 2.6 added -Recurse and Pre/PostCopyCommand options.
	Recurse works in conjunction with either delete option to remove folders.
	Pre/PostCopyCommand options come only from XML file.  They are strings
		that will be converted to a ScriptBlock and invoked on the target host.
#############################################################################>
