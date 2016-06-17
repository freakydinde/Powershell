<# 
.Synopsis
store user EMail credentials

.Description
store a file in data folder containing email credentials in CliXml format

.Example
Save-Credentials

.Notes  
    fileName	: Save-Credentials.ps1
    version		: 0.003
    author		: Armand Lacore
#>

$credentialsFolder = [IO.Path]::Combine($(Split-Path $PSScriptRoot), "Data", "Credentials")
if (!(Test-Path $credentialsFolder)) { New-Item $credentialsFolder -ItemType Directory -Force }

Get-Credential -Message "Please enter your Email credentials" | Export-Clixml ([IO.Path]::Combine($credentialsFolder, "mail@$($env:username)@$($env:computername).clixml"))
