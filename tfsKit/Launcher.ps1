<# 
.Synopsis
launch powershell, set global properties and import modules

.Notes  
    fileName		: Launcher.ps1
    version		   	: 0.04
    author         	: Armand Lacore
#>
Param (	[Parameter(Mandatory=$false,Position=0)][string]$LogsFolder,
		[Parameter(Mandatory=$false,Position=1)][string]$LogFile )

# reload that script with noExit, as an admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator") -and !$AdminNoExit)
{
	[array]$argumentsList = "-NoProfile", "-NoLogo", "-NoExit", "-File `"$($PSScriptRoot)\Launcher.ps1`""
	
	if ($LogsFolder) { $argumentsList += "-LogsFolder `"$LogsFolder`"" }
	if ($LogFile) { $argumentsList += "-LogFile `"$LogFile`"" }
	
	$argumentsList += "-AdminNoExit"
	
	Start-Process -File PowerShell.exe -Verb RunAs -ArgumentList $argumentsList
	Exit 0
}

# set $LogFile if passed as parameters
if ($LogFile) 
{ 
	$Global:LogFile = $LogFile
	$Global:LogsFolder = Split-Path $LogFile 
}
else
{
	# remove $LogFile if empty but existing
	if (Get-Variable LogFile -ErrorAction SilentlyContinue)
	{
		Remove-Variable LogFile -Force
	}

	# set $LogsFolder if passed as parameters
	if ($LogsFolder) 
	{ 
		$Global:LogsFolder = $LogsFolder 
	}
	else
	{
		# remove $LogsFolder if empty but existing
		if (Get-Variable LogsFolder -ErrorAction SilentlyContinue)
		{
			Remove-Variable LogsFolder -Force
		}
	}
}

# import Input\Ouput modules
if (!(Get-Module Module-IO)) { Import-Module ([IO.Path]::Combine($PSScriptRoot, "Modules", "Module-IO.psm1")) }

# Import all others modules
Add-Modules -Assert

# Error : Stop, Warning : Continue, Verbose : Continue, Debug : SilentlyContinue
Set-LogLevel "Verbose"

# Error : Red, Warning : Yellow, Debug : DarkGray, Verbose : Gray, Progress : Green
Set-LogColor

# set UI Size to large
Set-UISize

# set UI title
Set-UITitle "tfsKit"

# set PS current folder to rootFolder
Set-Location $global:RootFolder

Write-Host "`r`n`t tfsKit`r`n" -Fore Green
Write-Host @"
Get-TfsConnection`t `t `tGet-TfsCodeAnalysisInfo
Get-TfsBuilds`t `t `t `tGet-TfsCodeCoveragesInfo
Get-TfsBuildDetails`t `t `tGet-TfsCodeCoveragesInfo
Get-TfsLastExecutedBuildDetails`t `tGet-TfsCompilationInfo	
`t `t `t `t `tGet-TfsUnitTestsInfo
`r`n
"@ -Fore Cyan
Write-Host "`t Get-Help [FONCTION_NAME] -full`r`n" -Fore Green