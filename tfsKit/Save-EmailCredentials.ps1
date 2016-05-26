<# 
.Synopsis
store user EMail credentials

.Description
store a file in data folder containing email credentials in CliXml format

.Example
Save-Credentials

.Notes  
    fileName	: Save-Credentials.ps1
    version		: 0.002
    author		: Armand Lacore
#>

Get-Credential -Message "Please enter your Email credentials" | Export-Clixml ([IO.Path]::Combine($PSScriptRoot, "Data", "mail@$($env:username)@$($env:computername).clixml"))
