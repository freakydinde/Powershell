<# 
.Synopsis
launch powershell, set global properties and import modules

.Notes  
    fileName		: Launcher.ps1
    version		   	: 0.04
    author         	: Armand Lacore
#>

# reload that script with noExit
if ($args -NotContains "SecondLaunch")
{
	Start-Process PowerShell.exe -ArgumentList "-NoProfile -NoLogo -NoExit -File $($PSScriptRoot)\Launcher.ps1 SecondLaunch" 
    Exit 0
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