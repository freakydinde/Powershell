°Module-Tfs.psm1  

.Synopsis  
functionnality to query team foundation servers builds  

.Description  
contain method that perform operation on TFS projects collection, get builds data, compilation, code analysis & test results.  

.Notes   
	fileName	: Module-TFS.psm1  
	version		: 0.033  
	author		: Armand Lacore  

°TFS   

Get-TfsBuildDetails  
Get-TfsBuilds   
Get-TfsCodeAnalysisInfo  
Get-TfsCodeCoveragesInfo  
Get-TfsCompilationInfo  
Get-TfsConnection  
Get-TfsEnvironment  
Get-TfsLastExecutedBuildDetails  
Get-TfsUnitTestsInfo  
 
°Module-IO.psm1  

.Synopsis  
Input & Output helper  

.Description  
contain IO fonctionnality : log, xml, IO.  

.Notes  
	fileName	: Module-IO.psm1  
	version		: 0.52  
	author		: Armand Lacore  

°LOGGIN  

Add-ToEachLine  
Format-Message  
Set-UISize  
Set-LogColor  
Set-LogLevel  
Write-LogDebug  
Write-LogError  
Write-LogEvent  
Write-LogHost  
Write-LogObject  
Write-LogVerbose  
Write-LogWarning  
Write-ObjectToHtml  
Write-ObjectToXml  

°XML  

Get-Data  
Get-XPathValue  
Set-XPathValue  
Test-XPath  

°IO   

Add-Assemblies  
Add-Modules  
Assert-Folders  
Get-InString  
Get-Prompt  
Get-SecureString  
Get-SHA256  
Get-TempFolder  
Remove-Folder  
Set-SecureString  
