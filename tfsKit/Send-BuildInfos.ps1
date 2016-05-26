<# 
.Synopsis
Generate html builds report and send it through email 

.Description
Generate html builds report containing Compilation, Code Analysis, Unit Tests and code coverage informations, and send it through email getting data from BuildInfos.xml

.Parameter Language
Language used in code analysis short description (long description language depends on your tfs server settings)

.Parameter LogLevel
enable Host, Verbose, Debug or Trace log level

.Example
Send-BuildInfos

.Example
Send-BuildInfos "French"

.Example
Send-BuildInfos -LogLevel "Debug"

.Notes  
    fileName	: Send-BuildInfos.ps1
    version		: 0.019
    author		: Armand Lacore
#>

[CmdletBinding()] # enable -ErrorAction parameter
Param ( [Parameter(Mandatory=$false,Position=0)][Validateset("English","French","Spanish")][string]$Language="English",
  		[Parameter(Mandatory=$false,Position=1)][Validateset("Host","Verbose","Debug","Trace")][string]$LogLevel="Verbose" )

# import IO and TFS modules
$modulesShortName = @("IO", "Tfs")
$modulesShortName | if (!(Get-Module -Name "Module-$($_)")) { Import-Module ([IO.Path]::Combine($global:ModulesFolder, "Module-$($_).psm1")) -Force -Global }

# define verbose log level and log colors (functions from Module-IO)
Set-LogColor
Set-LogLevel $LogLevel

# create or clean temp folder
$tempFolder = [IO.Path]::Combine($global:RootFolder, "Temp")
if (!(Test-Path $tempFolder)) { New-Item $tempFolder -ItemType Directory -Force | Out-Null }
else { Get-ChildItem $tempFolder | % { Remove-Item $_.FullName -Force -ErrorAction Continue } }
		
try
{
	Write-LogDebug "Start Script Send-BuildInfos.ps1"

	Write-LogVerbose "Loading TFS configuration"

    # Load configuration from BuildInfos.xml
	$buildInfosData = [xml](Get-Content ([IO.Path]::Combine($gobal:DataFolder, "BuildInfos.xml")) -Encoding UTF8)
    $tfsCollectionUri = [string]$buildInfosData.BuildReports.Tfs.CollectionUri
    $tfsProjectName = [string]$buildInfosData.BuildReports.Tfs.ProjectName
	$tfsBuilds = $buildInfosData.BuildReports.Build

	# get tfs credentials if they exists (if not TFS will try to authenticate through active directory) 
	$tfsCredentialsPath = [IO.Path]::Combine($gobal:DataFolder, "tfs@$($env:username)@$($env:computername).clixml")
	if (Test-Path $tfsCredentialsPath) { $tfsCredentials = Import-Clixml $tfsCredentialsPath } else { $tfsCredentials = $null }

	Write-LogVerbose "Getting HTML template"

	# get and format html template
	$htmlCss = [string]$buildInfosData.BuildReports.Html.Css
	$htmlTemplate = [string]$buildInfosData.BuildReports.Html.Template
		# remove indentation inherited from xml template dirty temp todo check css template
	$htmlCss = $htmlCss -replace "`t`t`t", ""
	$htmlTemplate = $htmlTemplate -replace "`t`t`t", ""
		# add a proper 3 tab indent to CSS
    $htmlCss = Add-ToEachLine $htmlCss "`t`t`t"
		# add CSS to html template, remove blank lines, add 3 tab intend to first line (first line loose its intend while replacing)
	$htmlTemplate = ($htmlTemplate -replace '\(\$CSS\)', "`t`t`t$($htmlCss.Trim())").Trim()

	foreach ($tfsBuild in $tfsBuilds)
	{
		$tfsBuildName = $tfsBuild.name

		Write-LogHost "Querying $tfsBuildName"
		
		# get quality indicator treshold
		$codeAnalysisWarningsThreshold = [int]$tfsBuild.CodeAnalysisThreshold
		$codeCoverageWarningsThreshold = [int]$tfsBuild.CodeCoverageThreshold
        $compilationWarningsThreshold = [int]$tfsBuild.CompilationThreshold

		# get build data, tfsCredentials accept $null value, remove -RecentBuild switch if last build can be older than Yesterday (slow up processing)
        $build = Get-TfsLastExecutedBuildDetails $tfsCollectionUri $tfsProjectName $tfsBuildName $tfsCredentials $Language -FullDetails -RecentBuild

		Write-LogVerbose "Formating data from build $($build.BuildNumber)"

		# format build summary
		$buildDuration = "$([Math]::Round($($build.BuildDuration.TotalMinutes),2)) minuts"
		$codeAnalysisWarningsCount = [int]($build.CodeAnalysisInfo.Details | ? { $_.ErrorLevel -eq "Warning" }).Count
		$compilationWarningsCount = [int]($build.CompilationInfo.Details | ? { $_.ErrorLevel -eq "Warning" }).Count
		$testTotalCount = [int]($build.UnitTestsInfo.Details).Count
		$testCompletedCount = [int]($build.UnitTestsInfo.Details | ? { $_.Outcome -eq "Passed" }).Count
		$codeCoverage = [int]$build.CodeCoveragesInfo.Summary.Coverage
        if ($testTotalCount -gt 0) { $testCompletedPercentage = [Math]::Round(100 - (100 * ($testTotalCount - $testCompletedCount) / $testTotalCount),2) }
        else { $testCompletedPercentage = 0 }

        # format monitored values into red or green div depending on treshold values
        if ($build.CompilationSucceed -eq "True") { $compilationSucceed = "<div class=`"green`">$($build.CompilationSucceed)</div>" }
        else { $compilationSucceed = "<div class=`"red`">$($build.CompilationSucceed)</div>" }
        if ($build.CodeAnalysisSucceed -eq "True") { $codeAnalysisSucceed = "<div class=`"green`">$($build.CodeAnalysisSucceed)</div>" }
        else { $codeAnalysisSucceed = "<div class=`"red`">$($build.CodeAnalysisSucceed)</div>" }
        if ($testCompletedPercentage -eq 100) { $testCompletedPercentage = "<div class=`"green`">$testCompletedPercentage %</div>" } 
		else {  $testCompletedPercentage = "<div class=`"red`">$testCompletedPercentage %</div>" }
        if ($codeCoverage -ge $codeCoverageWarningsThreshold) { $codeCoverage = "<div class=`"green`">$codeCoverage %</div>" }
        else { $codeCoverage = "<div class=`"red`">$codeCoverage %</div>" }
        if ($compilationWarningsCount -le $compilationWarningsThreshold) { $compilationWarningsCount = "<div class=`"green`">$compilationWarningsCount</div>" }
        else { $compilationWarningsCount = "<div class=`"red`">$compilationWarningsCount</div>" }
        if ($codeAnalysisWarningsCount -le $codeAnalysisWarningsThreshold) { $codeAnalysisWarningsCount = "<div class=`"green`">$codeAnalysisWarningsCount</div>" }
        else { $codeAnalysisWarningsCount = "<div class=`"red`">$codeAnalysisWarningsCount</div>" }
	    
        # create build summary
        $buildSummary = [ordered]@{"Build name"=$($build.BuildNumber);"Status"=$($build.Status);"Date"=$($build.StartTime);"Duration"=$buildDuration;"Compilation succeed"=$compilationSucceed;"Code analysis succeed"=$codeAnalysisSucceed;"Test passed"=$testCompletedPercentage;"Code coverage"=$codeCoverage;"Compilation warnings"=$compilationWarningsCount;"Code analysis warnings"=$codeAnalysisWarningsCount }
        $buildSummaryHtml = Write-ObjectToHtml $buildSummary -TableClass "Summary" -TableId "bSummary" -Horizontal

		# get html tables from build data
		$codeAnalysisSummaryHtml = Write-ObjectToHtml $build.CodeAnalysisInfo.Summary -TableClass "Details" -TableId "caSummary"
		$codeAnalysisDetailsHtml = Write-ObjectToHtml $build.CodeAnalysisInfo.Details -TableClass "Details" -TableId "caDetails" -EncodeHtml
		$compilationDetailsHtml = Write-ObjectToHtml $build.CompilationInfo.Details -TableClass "Details" -TableId "compDetails" -EncodeHtml
		$unitTestSummaryHtml = Write-ObjectToHtml $build.UnitTestsInfo.Summary -TableClass "Summary" -TableId "utSum" -Horizontal
		$unitTestDetailsHtml = Write-ObjectToHtml $build.UnitTestsInfo.Details -TableClass "Details" -TableId "utDetails" -EncodeHtml
		$codeCoverageSummaryHtml = Write-ObjectToHtml $build.CodeCoveragesInfo.Summary -TableClass "Summary" -TableId "ccSum" -Horizontal
		$codeCoverageDetailsHtml = Write-ObjectToHtml $build.CodeCoveragesInfo.Details -TableClass "Details" -TableId "ccDetails" -EncodeHtml
        
		Write-LogVerbose "Creating email body"

        # build mail body
        $mailHtml = "<div class=`"mail`">`r`n"
        $mailHtml += "`t<h3 class=`"tableHeader`">Build summary</h3>`r`n"
        $mailHtml += "$(Add-ToEachLine $buildSummaryHtml "`t")"
        if (!$build.CompilationSucceed)
        {
            $compilationErrorFormated = @()
            $build.CompilationInfo.Details | ? { $_.ErrorLevel -eq "Error" } | % { $compilationErrorFormated += [ordered]@{Code=$_.Code; Message=$_.Message; File=$_.File; Line=$_.Line} }
            
            if ($compilationErrorFormated.Count -gt 0)
            {
                $compilationErrorHtml = Write-ObjectToHtml $compilationErrorFormated -TableClass "alert" -TableId "balert"

                $mailHtml += "<div class=`"alert`">`r`n"
			    $mailHtml += "`t<div class=`"alertTitle`">Compilation Errors</div>`r`n"
                $mailHtml += "$(Add-ToEachLine $compilationErrorHtml "`t")"
                $mailHtml += "</div>`r`n"
            }
        }
        if (!$build.CodeAnalysisSucceed)
        {
            $codeAnalysisErrorFormated = @()
            $build.CodeAnalysisInfo.Details | ? { $_.ErrorLevel -eq "Error" } | % { $codeAnalysisErrorFormated += [ordered]@{Code=$_.Code; Description=$_.Description; File=$_.File; Line=$_.Line} }

            if ($codeAnalysisErrorFormated.Count -gt 0)
            {
                $codeAnalysisErrorHtml = Write-ObjectToHtml $codeAnalysisErrorFormated -TableClass "alert" -TableId "balert"

                $mailHtml += "<div class=`"alert`">`r`n"
			    $mailHtml += "`t<div class=`"alertTitle`">Code Analysis Errors</div>`r`n"
                $mailHtml += "$(Add-ToEachLine $codeAnalysisErrorHtml "`t")"
                $mailHtml += "</div>`r`n"
            }
        }
        if ($codeAnalysisSummaryHtml)
        {
            $mailHtml += "`t<h3 class=`"tableHeader`">Code analysis</h3>`r`n"
            $mailHtml += "$(Add-ToEachLine $codeAnalysisSummaryHtml "`t")"
        }
        if ($unitTestSummaryHtml)
        {
            $mailHtml += "`t<h3 class=`"tableHeader`">Unit Tests</h3>`r`n"
            $mailHtml += "$(Add-ToEachLine $unitTestSummaryHtml "`t")"
        }
        if ($codeCoverageSummaryHtml)
        {
            $mailHtml += "`t<h3 class=`"tableHeader`">Code coverage</h3>`r`n"
            $mailHtml += "$(Add-ToEachLine $codeCoverageSummaryHtml "`t")"
        }
        $mailHtml += "</div>"
        # add a 2 tab indentation to mail body before inserting into Html template
        $mailHtml = Add-ToEachLine $mailHtml "`t`t"
        # insert mail html into html template, add 2 tab intend to first line (first line loose its intend while replacing)
        $mailHtml = $htmlTemplate -replace '\(\$Title\)', "build report" -replace '\(\$Body\)', "`t`t$($mailHtml.Trim())"

		Write-LogVerbose "Creating email attachments files"
        
		# build attachments files        
        $attachments = @()
        if ($compilationDetailsHtml)
        {
            $htmlPage = $htmlTemplate -replace '\(\$Title\)',"Compilation"
		    $htmlPage = $htmlPage -replace '\(\$Body\)', $(Add-ToEachLine $compilationDetailsHtml "`t`t").Trim()
		    $htmlPage | Out-File "$tempFolder\Compilation.Html"
            $attachments += "$tempFolder\Compilation.Html"
        }
        if ($unitTestDetailsHtml)
        {
            $htmlPage = $htmlTemplate -replace '\(\$Title\)',"UnitTests"
		    $htmlPage = $htmlPage -replace '\(\$Body\)', $(Add-ToEachLine $unitTestDetailsHtml "`t`t").Trim()
		    $htmlPage | Out-File "$tempFolder\UnitTests.Html"
            $attachments += "$tempFolder\UnitTests.Html"
        }
        if ($codeAnalysisDetailsHtml)
        {
            $htmlPage = $htmlTemplate -replace '\(\$Title\)',"CodeAnalysis"
		    $htmlPage = $htmlPage -replace '\(\$Body\)', $(Add-ToEachLine $codeAnalysisDetailsHtml "`t`t").Trim()
		    $htmlPage | Out-File "$tempFolder\CodeAnalysis.Html"
            $attachments += "$tempFolder\CodeAnalysis.Html"
        }
        if ($codeCoverageDetailsHtml)
        {
            $htmlPage = $htmlTemplate -replace '\(\$Title\)',"CodeCoverage"
		    $htmlPage = $htmlPage -replace '\(\$Body\)', $(Add-ToEachLine $codeCoverageDetailsHtml "`t`t").Trim()
		    $htmlPage | Out-File "$tempFolder\CodeCoverage.Html"
            $attachments += "$tempFolder\CodeCoverage.Html"
        }

		Write-LogVerbose "Getting email configuration"

		# get email credentials
        $emailCredentialsPath = [IO.Path]::Combine($global:DataFolder, "Credentials", "mail@$($env:USERNAME)@$($env:COMPUTERNAME).clixml")
		if (Test-Path $emailCredentialsPath) { $emailCredentials = Import-Clixml $emailCredentialsPath }
		else { Throw "credentials missing, email won't be send" }
        
		# get email title and recipients
        $subject = "Build report $($build.BuildNumber)"
		$recipients = [string[]]$tfsBuild.Recipients.SendTo
        
		Write-LogVerbose "Sending email $subject to $recipients"

        # i use System.Net.Mail.SmtpClient Object instead of Send-MailMessage function to get more controls and dispose functionnality
        try
        {
			# create mail and server objects
            $message = New-Object -TypeName System.Net.Mail.MailMessage
            $smtp = New-Object -TypeName System.Net.Mail.SmtpClient($buildInfosData.BuildReports.Mail.Server)

			# build message
            $recipients | % { $message.To.Add($_) }
            $message.Subject = $subject
            $message.From = New-Object System.Net.Mail.MailAddress($emailCredentials.UserName)
            $message.Body = $mailHtml
            $message.IsBodyHtml = $true
			$attachments | % { $message.Attachments.Add($(New-Object System.Net.Mail.Attachment $_)) }
			
			# build SMTP server
            $smtp = New-Object -TypeName System.Net.Mail.SmtpClient([string]$buildInfosData.BuildReports.Mail.Server)
            $smtp.Port = [int]$buildInfosData.BuildReports.Mail.Port
            $smtp.Credentials = [System.Net.ICredentialsByHost]$emailCredentials
            $smtp.EnableSsl = [bool]$buildInfosData.BuildReports.Mail.UseSsl

			# send message
            $smtp.Send($message)

			Write-LogHost "Email message sent" 
        }
        catch
        {
            Write-LogWarning "$($_.Exception | Select Message, Source, ErrorCode, InnerException, StackTrace | Format-List | Out-String)" 
        }
        finally
        {
            Write-LogVerbose "Disposing Smtp Object"
            $message.Dispose()
            $smtp.Dispose()
        }

		Write-LogDebug "Success Script Send-BuildInfos.ps1"
	}
}
catch
{
    Write-LogError $($_.Exception) $($_.InvocationInfo)
}
finally
{
	# remove temp folder
	Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction Continue 
}
