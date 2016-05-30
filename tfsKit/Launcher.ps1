<# 
.Synopsis
launch powershell, set global properties and import modules

.Notes  
    fileName		: Launcher.ps1
    version		   	: 0.02
    author         	: Armand Lacore
#>

# reload that script with noExit
if ($args -NotContains "SecondLaunch")
{
	Start-Process PowerShell.exe -ArgumentList "-NoProfile -NoLogo -NoExit -File $($PSScriptRoot)\Launcher.ps1 SecondLaunch" 
    Exit 0
}

# Import all kit modules
Get-ChildItem ([IO.Path]::Combine($PSScriptRoot, "Modules")) | % { Import-Module $_.FullName -Force -Global }

# Set-LogLevel, Verbose = Error : Stop, Warning : Continue, Verbose : Continue, Debug : SilentlyContinue
Set-LogLevel "Verbose"

# Set-LogColor, default = Error : Red, Warning : Yellow, Debug : DarkGray, Verbose : Gray, Progress : Green
Set-LogColor

# Set-UISize, default = Width : 150, Height : 50
Set-UISize

# set current folder to rootFolder
Set-Location $global:RootFolder