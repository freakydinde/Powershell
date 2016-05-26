<# 
.Synopsis
team foundation server Functionnality

.Description
contain method that perform operation on tfs projects collection

.Notes 
	fileName	: Module-TFS.psm1
	version		: 0.013
	author		: Armand Lacore
#>

#region <MODULE PARAMETERS>

# import Input\Output module
if (!(Get-Module -Name Module-IO)) { Import-Module ([IO.Path]::Combine($PSScriptRoot,"Module-IO.psm1")) }

# import TFS assemblies ($global:AssembliesFolder is set on Module-IO.psm1 import)
$assembliesShortName = @("TeamFoundation.Build.Common","TeamFoundation.Client","TeamFoundation.TestManagement.Client","TeamFoundation.WorkItemTracking.Client","TeamFoundation.WorkItemTracking.Client.DataStoreLoader","VisualStudio.Services.Client")
$assembliesShortName | Add-Type -Path $([IO.Path]::Combine($global:AssembliesFolder, "Microsoft.$($_).dll")) 

#endregion

#region <MODULE FUNCTIONS>

<# 
.Synopsis
Return TFS builds informations

.Description
Get build details from build number, optionally add full detailled informations

.Parameter TfsCollectionUri
Tfs Team project Uri

.Parameter TfsProjectName
Project Name from TFS collection

.Parameter TfsBuildNumber
build number, including buildName_buildDate

.Parameter Language
Language used in code analysis error descriptions (only used in FullDetails mode)

.Parameter Details
switch, set to true if you wan't to get build errors, build warnings, UnitTest and code coverage results

.Parameter FullDetails
switch, set to true if you wan't to get build errors, build warnings, UnitTest and code coverage results and full detailled informations array

.Example
Get-TfsBuildDetails "https://tfs.axacolor.igo6.com/tfs/DefaultCollection" "CI-IGO6-Main" "CI-IGO6-Main_20160306.1" -FullDetails

.Outputs
Tfs build details on success / null on fail
#> 
Function Get-TfsBuildDetails
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][string]$TfsCollectionUri,
			[Parameter(Mandatory=$true,Position=1)][string]$TfsProjectName,
			[Parameter(Mandatory=$true,Position=2)][string]$TfsBuildNumber,
			[Parameter(Mandatory=$false,Position=3)][Validateset("English","French","Spanish")][string]$Language="English",
			[Parameter(Mandatory=$false,Position=4)][switch]$Details,
			[Parameter(Mandatory=$false,Position=5)][switch]$FullDetails )

    Write-LogDebug "Start Get-TfsBuildDetails"

    try
    {
        Write-LogHost "getting build details from $tfsBuildNumber" -LineBefore

        $tfsBuildName = $tfsBuildNumber -replace '^(.*)_(\d{8}?)\.(\d*)$', '$1'

        $tfsBuilds = Get-TfsBuilds $TfsCollectionUri $TfsProjectName $tfsBuildName -BuildNumber $tfsBuildNumber -InformationTypes "*"
      
		if ($tfsBuilds)
		{
			$tfsBuild = $tfsBuilds.Builds | Select-Object -First 1
			$tfsBuild | Add-Member -NotePropertyName Tfs -NotePropertyValue $($tfsBuilds.Tfs)

			if ($Details -or $FullDetails)
			{
				$tfsBuild | Add-Member -NotePropertyName BuildDuration -NotePropertyValue $($tfsBuild.FinishTime - $tfsBuild.StartTime)

				# get compilation and code analysis errors and warnings
				Write-LogVerbose "getting build errors and warning"
				$buildErrors = [Microsoft.TeamFoundation.Build.Client.InformationNodeConverters]::GetBuildErrors($tfsBuild)               
				$buildWarnings = [Microsoft.TeamFoundation.Build.Client.InformationNodeConverters]::GetBuildWarnings($tfsBuild)

				$compilationErrors = $buildErrors | ? { $_.ErrorType -eq "Compilation" }
				$compilationWarnings = $buildWarnings | ? { $_.WarningType -eq "Compilation" } 
				$codeAnalysisErrors = $buildErrors | ? { $_.ErrorType -eq "StaticAnalysis" }
				$codeAnalysisWarnings = $buildWarnings | ? { $_.WarningType -eq "StaticAnalysis" } 

				# get compilation info
				Write-LogVerbose "getting build compilation results"
				$compilation = @()
				foreach ($compilationMessage in $compilationErrors)
				{
					$compilationMessage | Add-Member -NotePropertyName ErrorLevel -NotePropertyValue "Error"
					$compilation += $compilationMessage
				}
				foreach ($compilationMessage in $compilationWarnings)
				{
					$compilationMessage | Add-Member -NotePropertyName ErrorLevel -NotePropertyValue "Warning"
					$compilation += $compilationMessage
				}
				$compilationIsSuccessFull = $(($compilation | ? { $_.ErrorLevel -eq "Error" }).Count -le 0)
				$tfsBuild | Add-Member -NotePropertyName CompilationSucceed -NotePropertyValue $compilationIsSuccessFull
				$tfsBuild | Add-Member -NotePropertyName Compilation -NotePropertyValue $compilation

				# get code analysis info
				Write-LogVerbose "getting build code analysis results"
				$codeAnalysis = @()
				foreach ($codeAnalysisMessage in $codeAnalysisErrors)
				{
					$codeAnalysisMessage | Add-Member -NotePropertyName ErrorLevel -NotePropertyValue "Error"
					$codeAnalysis += $codeAnalysisMessage
				}
				foreach ($codeAnalysisMessage in $codeAnalysisWarnings)
				{
					$codeAnalysisMessage | Add-Member -NotePropertyName ErrorLevel -NotePropertyValue "Warning"
					$codeAnalysis += $codeAnalysisMessage
				}    
				$codeAnalysisIsSuccessFull = $(($codeAnalysis | ? { $_.ErrorLevel -eq "Error" }).Count -le 0)
				$tfsBuild | Add-Member -NotePropertyName CodeAnalysisSucceed -NotePropertyValue $codeAnalysisIsSuccessFull
				$tfsBuild | Add-Member -NotePropertyName CodeAnalysis -NotePropertyValue $codeAnalysis

				# get unit tests info
				Write-LogVerbose "getting build unit tests results"
				$unitTests = $tfsBuild.Tfs.TestManagementService.TestRuns.ByBuild($tfsBuild.Uri)

				$testAreSuccessFull = $true
				foreach ($testsSet in $tfsBuild.UnitTest)
				{
					if ([int]($testsSet.Statistics.FailedTests) -gt 0) { $testAreSuccessFull = $false }
				}
				$tfsBuild | Add-Member -NotePropertyName UnitTestsSucceed -NotePropertyValue $testAreSuccessFull
				$tfsBuild | Add-Member -NotePropertyName UnitTests -NotePropertyValue $unitTests

				Write-LogVerbose "getting build code coverage"
				$codeCoverages = $tfsBuild.Tfs.TestManagementService.CoverageAnalysisManager.QueryBuildCoverage($tfsBuild.Uri,[Microsoft.TeamFoundation.TestManagement.Client.CoverageQueryFlags]::BlockData -bor [Microsoft.TeamFoundation.TestManagement.Client.CoverageQueryFlags]::Functions -bor [Microsoft.TeamFoundation.TestManagement.Client.CoverageQueryFlags]::Modules)
				
				$tfsBuild | Add-Member -NotePropertyName CodeCoverages -NotePropertyValue $($codeCoverages)
			}

			if ($FullDetails)
			{
				$tfsBuild | Add-Member -NotePropertyName CompilationInfo -NotePropertyValue $(Get-TfsCompilationInfo $tfsBuild)
				$tfsBuild | Add-Member -NotePropertyName CodeAnalysisInfo -NotePropertyValue $(Get-TfsCodeAnalysisInfo $tfsBuild -Language $Language)
				$tfsBuild | Add-Member -NotePropertyName UnitTestsInfo -NotePropertyValue $(Get-TfsUnitTestsInfo $tfsBuild)
				$tfsBuild | Add-Member -NotePropertyName CodeCoveragesInfo -NotePropertyValue $(Get-TfsCodeCoveragesInfo $tfsBuild)
			}
		}
		else
		{
			Throw "no build data for $tfsBuildNumber"
		}
		
	    # trace Success
	    Write-LogDebug "Success Get-TfsBuildDetails"
    }
    catch
    {
        $tfsBuild = $null

        # log Error 
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $tfsBuild
}  

<# 
.Synopsis
Return TFS builds informations

.Description
Use build details specification properties to filter type of data returned,see https://msdn.microsoft.com/en-us/library/microsoft.teamfoundation.build.server.builddetailspec_properties(v=vs.120).aspx

.Parameter TfsCollectionUri
Tfs Team project Uri

.Parameter TfsProjectName
Project Name from TFS collection

.Parameter TfsBuildName
Name of the builds to return

.Parameter TfsBuildNumber
build number filter, wildcards are allowed, default = "*"

.Parameter InformationTypes
information types to return, * indicates all informations, default = $null

.Parameter MinFinishTime
minimum finish time filter, default = 0 (01/01/0001 00:00:00)

.Parameter MaxFinishTime
maximum finish time filter, default = 0 (01/01/0001 00:00:00)

.Parameter QueryOptions
options that are used to control the amount of data returned, default = All, available : None, Definitions, Agents, Workspaces, Controllers, Process, BatchedRequests, HistoricalBuilds, All

.Parameter QueryOrder
desired order of the builds returned, default = StartTimeAscending, available : StartTimeAscending, StartTimeDescending, FinishTimeAscending, FinishTimeDescending

.Parameter Status
build status filter, default = All, available : None, InProgress, Succeeded, PartiallySucceeded, Failed, Stopped, NotStarted, All

.Example
Get-TfsBuilds "https://tfs.axacolor.igo6.com/tfs/DefaultCollection" "COLOR-SP" "CI-IGO6-Main"

.Outputs
Tfs builds on success / null on fail
#> 
Function Get-TfsBuilds 
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][string]$TfsCollectionUri,
			[Parameter(Mandatory=$true,Position=1)][string]$TfsProjectName,
			[Parameter(Mandatory=$true,Position=2)][string]$TfsBuildName,
			[Parameter(Mandatory=$false,Position=3)][string]$BuildNumber="*",
			[Parameter(Mandatory=$false,Position=4)][string]$InformationTypes=$null,
			[Parameter(Mandatory=$false,Position=5)][DateTime]$MinFinishTime=0,
			[Parameter(Mandatory=$false,Position=6)][DateTime]$MaxFinishTime=0,
			[Parameter(Mandatory=$false,Position=7)][ValidateSet("None","Definitions","Agents","Workspaces","Controllers","Process","BatchedRequests","HistoricalBuilds","All")][string]$QueryOptions="All",
			[Parameter(Mandatory=$false,Position=8)][ValidateSet("StartTimeAscending","StartTimeDescending","FinishTimeAscending","FinishTimeDescending")][string]$QueryOrder="StartTimeAscending",
			[Parameter(Mandatory=$false,Position=9)][ValidateSet("None","InProgress","Succeeded","PartiallySucceeded","Failed","Stopped","NotStarted","All")][string]$Status="All" )

    Write-LogDebug "Start Get-TfsBuilds"

    try
    {
        $tfs = Get-TfsConnection $TfsCollectionUri

	    if ($tfs.HasAuthenticated)
	    {
            # create build specification
            $buildSpecification = $tfs.BuildServer.CreateBuildDetailSpec($TfsProjectName, $TfsBuildName)
            $buildSpecification.BuildNumber = $BuildNumber
            $buildSpecification.InformationTypes = $InformationTypes
            $buildSpecification.MinFinishTime = $MinFinishTime
            $buildSpecification.MaxFinishTime = $MaxFinishTime
            $buildSpecification.QueryOptions = $QueryOptions
            $buildSpecification.QueryOrder = $QueryOrder
            $buildSpecification.Status = $Status

			Write-LogObject $buildSpecification "getting build with following specification" -LineAfterTitle

			Write-LogVerbose "getting TFS build $TfsBuildName"

			$tfsBuilds = $tfs.BuildServer.QueryBuilds($buildSpecification)
			
			if ($tfsBuilds)
			{
				$tfsBuilds | Add-Member -NotePropertyName Tfs -NotePropertyValue $tfs
			}
			else
			{
				Throw "no build data for build name $TfsBuildName & build number : $BuildNumber"	
			}
        }
        else
        {
            Throw "you must be authenticated to TFS projects collection to get builds data"
        }

	    # trace Success
	    Write-LogDebug "Success Get-TfsBuilds"
    }
    catch
    {
        $tfsBuilds = $null

        # log Error 
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $tfsBuilds
} 

<# 
.Synopsis
Return code analysis details from tfs builds

.Description
return an array containing code analysis errors and warning occurency by ids

.Parameter TfsBuild
full build containing all build data

.Parameter Language
Language used in code analysis error descriptions

.Example
Get-TfsCodeAnalysisInfo $tfsBuild

.Outputs
code analysis details on success / null on fail
#> 
Function Get-TfsCodeAnalysisInfo
{
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true,Position=0)][object]$TfsBuild,
			[Parameter(Mandatory=$false,Position=1)][Validateset("English","French","Spanish")][string]$Language="English" )

    Write-LogDebug "Start Get-TfsCodeAnalysisInfo"

    try
    {
        Write-LogHost "getting code analysis info from $($TfsBuild.BuildNumber)"

        # Load configuration
        $buildInfoParameters = Get-Data "BuildInfos"

        # get code analysis rules code mapping containing translation
        $codeAnalysisRules = $buildInfoParameters.BuildReports.CodeAnalysisRules.CodeAnalysisRule

		# initialize returned object
		$codeAnalysisInfo = New-Object -TypeName PsObject
		$summaryArray = @()
		$detailsArray = @()

		# build summary
        $codeAnalysisMessageGroup = $TfsBuild.CodeAnalysis | Group-Object Code | Sort-Object Count -Descending

        foreach ($codeAnalysisMessage in $codeAnalysisMessageGroup)
        {
            $codeDescription = $codeAnalysisRules | ? { $_.ID -eq $($codeAnalysisMessage.Name) } | Select-Object -First 1
            $summaryArray += $([ordered]@{Code=$($codeDescription.ID);Description=$($codeDescription.$Language);Count=$($codeAnalysisMessage.Count) })
        }

		$codeAnalysisInfo | Add-Member -NotePropertyName Summary -NotePropertyValue $summaryArray

		# build detailled array
		foreach ($codeAnalysisMessage in $($TfsBuild.CodeAnalysis | Sort-Object Code))
        {
            $codeDescription = $codeAnalysisRules | ? { $_.ID -eq $($codeAnalysisMessage.Code) } | Select-Object -First 1
            $detailsArray += $([ordered]@{Code=$($codeDescription.ID);Description=$($codeDescription.$Language);Message=$($codeAnalysisMessage.Message);File=$($codeAnalysisMessage.File);Line=$($codeAnalysisMessage.LineNumber);Link=$($codeDescription.Link);ErrorLevel=$($codeAnalysisMessage.ErrorLevel) })
        }

		$codeAnalysisInfo | Add-Member -NotePropertyName Details -NotePropertyValue $detailsArray

	    # trace Success
	    Write-LogDebug "Success Get-TfsCodeAnalysisInfo"
    }
    catch
    {
        $codeAnalysisInfo = $null

        # log Error 
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $codeAnalysisInfo
} 

<# 
.Synopsis
Return code coverage details from tfs builds

.Description
return an array containing code coverage errors and warning occurency by ids

.Parameter TfsBuild
full build containing all build data

.Example
Get-TfsCodeCoveragesInfo $tfsBuild

.Outputs
code coverage details on success / null on fail
#> 
Function Get-TfsCodeCoveragesInfo
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][object]$TfsBuild )

    Write-LogDebug "Start Get-TfsCodeCoveragesInfo"

    try
    {
        Write-LogHost "getting code coverage from $($TfsBuild.BuildNumber)"

		# initialize returned object
		$returnArray = @()

		foreach ($codeCoverage in $TfsBuild.CodeCoverages)
		{
			$codeCoveragesArray = New-Object -TypeName PsObject
			$codeCoveragesArray | Add-Member -NotePropertyName ID -NotePropertyValue $codeCoverage.Configuration.Id

			$modules = @()
			foreach ($module in $codeCoverage.Modules)
			{
				$moduleTotalLines = $module.Statistics.LinesCovered + $module.Statistics.LinesNotCovered + $module.Statistics.LinesPartiallyCovered
				$moduleTotalBlocks = $module.Statistics.BlocksNotCovered + $module.Statistics.BlocksCovered
				$moduleCoverage = [Math]::Round($([double]($module.Statistics.BlocksCovered) / $moduleTotalBlocks * 100),5)
                                      
				$modules += [ordered]@{Module=$module.Name;Coverage=$moduleCoverage;TotalBlocks=$moduleTotalBlocks;BlocksCovered=$module.Statistics.BlocksCovered;BlocksNotCovered=$module.Statistics.BlocksNotCovered;TotalLines=$moduleTotalLines;LinesCovered=$module.Statistics.LinesCovered;LinesPartiallyCovered=$module.Statistics.LinesPartiallyCovered;LinesNotCovered=$module.Statistics.LinesNotCovered}
			}

			$blocksCovered = $($modules.BlocksCovered | Measure-Object -Sum).Sum
			$blocksNotCovered = $($modules.BlocksNotCovered | Measure-Object -Sum).Sum
			$totalBlocks = $($modules.TotalBlocks | Measure-Object -Sum).Sum
			$linesCovered = $($modules.LinesCovered | Measure-Object -Sum).Sum
			$linesNotCovered = $($modules.LinesNotCovered | Measure-Object -Sum).Sum
			$linesPartiallyCovered = $($modules.PartiallyCovered | Measure-Object -Sum).Sum
			$totalLines = $($modules.TotalLines | Measure-Object -Sum).Sum
            
            if ($totalBlocks -gt 0)
            {
                $coverage = [Math]::Round($([double]$blocksCovered / $totalBlocks * 100),2)
            }
            else
            {
                $coverage = 0
            }

			$statistics = [ordered]@{BlocksCovered=$blocksCovered;BlocksNotCovered=$blocksNotCovered;TotalBlocks=$totalBlocks;LinesCovered=$linesCovered;LinesNotCovered=$linesNotCovered;LinesPartiallyCovered=$linesPartiallyCovered;TotalLines=$totalLines;Coverage=$coverage}

			$codeCoveragesArray | Add-Member -NotePropertyName Details -NotePropertyValue $modules
			$codeCoveragesArray | Add-Member -NotePropertyName Summary -NotePropertyValue $statistics

			$returnArray += $CodeCoveragesArray
		}

	    # trace Success
	    Write-LogDebug "Success Get-TfsCodeCoveragesInfo"
    }
    catch
    {
        $returnArray = $null

        # log Error 
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $returnArray
} 

<# 
.Synopsis
Return build messages details from tfs builds

.Description
return an array containing build messages details

.Parameter TfsBuild
full build containing all build data

.Example
Get-TfsCompilationInfo $tfsBuild

.Outputs
Build messages details on success / null on fail
#> 
Function Get-TfsCompilationInfo
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][object]$TfsBuild,
			[Parameter(Mandatory=$false,Position=1)][string]$Language="English")

    Write-LogDebug "Start Get-TfsCompilationInfo"

    try
    {
        Write-LogHost "getting compilation info from $($TfsBuild.BuildNumber)"

        # Load configuration
        $buildInfoData = Get-Data "BuildInfos"

		# initialize returned object
		$compilationInfo = New-Object -TypeName PsObject
		$summaryArray = @()
		$detailsArray = @()

		# build summary
        $compilationMessageGroup = $TfsBuild.Compilation | Group-Object Code | Sort-Object Count -Descending

        foreach ($compilationMessage in $compilationMessageGroup)
        {
            $summaryArray += $([ordered]@{Code=$($compilationMessage.Name);Count=$($compilationMessage.Count) })
        }

		$compilationInfo | Add-Member -NotePropertyName Summary -NotePropertyValue $summaryArray

		# build detailled array
		foreach ($compilationMessage in $($TfsBuild.Compilation))
        {
            $detailsArray += $([ordered]@{Code=$($compilationMessage.Code);Message=$($compilationMessage.Message);File=$($compilationMessage.File);Line=$($compilationMessage.LineNumber);ErrorLevel=$($compilationMessage.ErrorLevel) })
        }

		$compilationInfo | Add-Member -NotePropertyName Details -NotePropertyValue $detailsArray

	    # trace Success
	    Write-LogDebug "Success Get-TfsCompilationInfo"
    }
    catch
    {
        $compilationInfo = $null

        # log Error 
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $compilationInfo
} 

<# 
.Synopsis
Get Tfs Collection and server from Project Uri

.Description
Use GetTeamProjectCollection method from TfsTeamProjectCollectionFactory class to get Project Collection from Project Uri, check if Tfs is authenticated, if so get builds server, work item and test management service. 

.Parameter TfsCollectionUri
Tfs Team project Uri

.Example
Get-TfsConnection "https://tfs.axacolor.igo6.com/tfs/DefaultCollection"

.Outputs
TfsCollection on success / null on fail
#> 
Function Get-TfsConnection
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][System.Uri]$TfsCollectionUri )

    Write-LogDebug "Start Get-TfsConnection"

    try
    {
        Write-LogHost "getting TFS build server from $TfsCollectionUri" -LineBefore

		# initializing return object
		$tfsConnection = New-Object -TypeName PSObject

        Write-LogHost "getting TFS collection"
		
		$projectsCollection = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($TfsCollectionUri)
		
        Write-LogVerbose "checking authentication to Team projects collection"
		$projectsCollection.EnsureAuthenticated()

        $tfsConnection | Add-Member -NotePropertyName "HasAuthenticated" -NotePropertyValue ($projectsCollection.HasAuthenticated)

		if($projectsCollection.HasAuthenticated)
		{
            Write-LogVerbose "user is authenticated"

			Write-LogVerbose "getting TFS builds server from projects collection $projectsCollection"
			$buildServer = $projectsCollection.GetService([Microsoft.TeamFoundation.Build.Client.IBuildServer])
			$tfsConnection | Add-Member -NotePropertyName "BuildServer" -NotePropertyValue $buildServer

			Write-LogVerbose "getting TFS work item store cache from projects collection $projectsCollection"
			$workItemStore = $projectsCollection.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
			$tfsConnection | Add-Member -NotePropertyName "WorkItemStore" -NotePropertyValue $workItemStore

			Write-LogVerbose "getting test management services from projects collection $projectsCollection"
			$testManagementService = $projectsCollection.GetService([Microsoft.TeamFoundation.TestManagement.Client.ITestManagementService])
			$tfsConnection | Add-Member -NotePropertyName "TestManagementService" -NotePropertyValue $testManagementService
		}
		else
		{
			Write-LogWarning "your are not authenticated to $TfsCollectionUri"
		}

	    # trace Success
	    Write-LogDebug "Success Get-TfsConnection"
    }
    catch
    {
        $tfsConnection = $null

        # log Error 
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $tfsConnection
}   

<#
.Synopsis
Get TFS environment variables

.Description
get tfs environment serialized into PsObject,work only with TFS builds

.Example
Get-TfsEnvironment

.Outputs
PSObject containing TFS env,null on fail
#>
Function Get-TfsEnvironment
{
	try
	{
		$userName = "$($env:USERDOMAIN)\$($env:USERNAME)"
		$date = Get-Date -Format "yyyy.MM.dd_HH.mm.ss.fff"

		$tfs = New-Object -TypeName PSObject
		$tfs | Add-Member -NotePropertyName User -NotePropertyValue $userName
		$tfs | Add-Member -NotePropertyName DateTime -NotePropertyValue $date
		$tfs | Add-Member -NotePropertyName BuildName -NotePropertyValue $env:TF_BUILD_BUILDDEFINITIONNAME
		$tfs | Add-Member -NotePropertyName BuildDirectory -NotePropertyValue $env:TF_BUILD_BUILDDIRECTORY
		$tfs | Add-Member -NotePropertyName BuildReason -NotePropertyValue $env:TF_BUILD_BUILDREASON
		$tfs | Add-Member -NotePropertyName BuildUri -NotePropertyValue $env:TF_BUILD_BUILDURI
		$tfs | Add-Member -NotePropertyName CollectionUri -NotePropertyValue $env:TF_BUILD_COLLECTIONURI
		$tfs | Add-Member -NotePropertyName DropLocation -NotePropertyValue $env:TF_BUILD_DROPLOCATION
		$tfs | Add-Member -NotePropertyName Changeset -NotePropertyValue ($env:TF_BUILD_SOURCEGETVERSION -replace '^C(.*)$','$1')
		$tfs | Add-Member -NotePropertyName SourcesDirectory -NotePropertyValue $env:TF_BUILD_SOURCESDIRECTORY
		$tfs | Add-Member -NotePropertyName TestResultDirectory -NotePropertyValue $env:TF_BUILD_TESTRESULTSDIRECTORY

		Write-LogObject $tfs "TFS environments variables"
	}
	catch
	{
		# log warning
		Write-LogWarning "Error Get-TfsEnvironment -Exception $($_.Exception.Message)"
		$tfs = $null
	}

	return $tfs
}

<# 
.Synopsis
Return last executed TFS builds informations

.Description
First perfom a fast search on all builds definition to get last TfsBuildNumber, then a deep search on that build number

.Parameter TfsCollectionUri
Tfs Team project Uri
 
.Parameter TfsProjectName
Project Name from TFS collection

.Parameter TfsBuildName
Name of the builds to return

.Parameter Language
Language used in code analysis error descriptions (only used in FullDetails mode)

.Parameter Details
switch, set to true if you wan't to get build errors, build warnings, UnitTest and code coverage results

.Parameter FullDetails
switch, set to true if you wan't to get build errors, build warnings, UnitTest and code coverage results and full detailled informations array

.Parameter RecentBuild
switch, set to true if your build is less than 24h age, to speed up build search

.Example
Get-TfsLastExecutedBuildDetails "https://tfs.axacolor.igo6.com/tfs/DefaultCollection" "COLOR-SP" "CI-IGO6-Main"

.Outputs
Tfs build on success / null on fail
#> 
Function Get-TfsLastExecutedBuildDetails
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][string]$TfsCollectionUri,
			[Parameter(Mandatory=$true,Position=1)][string]$TfsProjectName,
			[Parameter(Mandatory=$true,Position=2)][string]$TfsBuildName,
			[Parameter(Mandatory=$false,Position=4)][Validateset("English","French","Spanish")][string]$Language="English",
			[Parameter(Mandatory=$false,Position=5)][switch]$Details,
			[Parameter(Mandatory=$false,Position=6)][switch]$FullDetails,
			[Parameter(Mandatory=$false,Position=7)][switch]$RecentBuild )

    Write-LogDebug "Start Get-TfsLastExecutedBuildDetails"

    try
    {
		if ($RecentBuild)
		{
			$yesterday = (Get-Date).AddDays(-1)
			$tfsBuilds = Get-TfsBuilds $TfsCollectionUri $TfsProjectName $TfsBuildName -QueryOrder "StartTimeDescending" -QueryOptions "Definitions" -MinFinishTime $yesterday
		}
		else
		{
			$tfsBuilds = Get-TfsBuilds $TfsCollectionUri $TfsProjectName $TfsBuildName -QueryOrder "StartTimeDescending" -QueryOptions "Definitions"
		}

        $tfsLastBuildNumber = ($tfsBuilds.Builds | ? { $_.Status -ne "InProgress" } | Select-Object -First 1).BuildNumber

        $tfsBuildDetails = Get-TfsBuildDetails $TfsCollectionUri $TfsProjectName $tfsLastBuildNumber $Language $Details $FullDetails

	    # trace Successs
	    Write-LogDebug "Success Get-TfsLastExecutedBuildDetails"
    }
    catch
    {
        $tfsBuildDetails = $null

        # log Error 
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $tfsBuildDetails
} 

<# 
.Synopsis
Return code analysis details from tfs builds

.Description
return an array containing code analysis errors and warning occurency by ids

.Parameter TfsBuild
full build containing all build data

.Example
Get-TfsUnitTestsInfo $tfsBuild

.Outputs
code analysis details on success / null on fail
#> 
Function Get-TfsUnitTestsInfo
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][object]$TfsBuild )

    Write-LogDebug "Start Get-TfsUnitTestsInfo"

    try
    {
        Write-LogHost "getting unit tests info from $($TfsBuild.BuildNumber)"

		# initialize returned object
		$returnArray = @()

		foreach ($testsSet in $TfsBuild.UnitTests)
		{
			$unitTestArray = New-Object -TypeName PsObject
			$unitTestArray | Add-Member -NotePropertyName Title -NotePropertyValue $testsSet.Title

			$statisticsObject = [ordered]@{Total=$($testsSet.Statistics.TotalTests);Completed=$($testsSet.Statistics.CompletedTests);Passed=$($testsSet.Statistics.PassedTests);Failed=$($testsSet.Statistics.FailedTests);Inconclusive=$($testsSet.Statistics.InconclusiveTests) }
			$unitTestArray | Add-Member -NotePropertyName Summary -NotePropertyValue $statisticsObject

			$detailsArray = @()
			foreach ($testResult in $testsSet.QueryResults())
			{
				$detailsArray += $([ordered]@{Title=$($testResult.TestCaseTitle);State=$($testResult.State);Outcome=$($testResult.Outcome);"Duration (ms)"=$($testResult.Duration.TotalMilliseconds);Storage=$($testResult.Implementation.Storage);ErrorMessage=$($testResult.ErrorMessage) })
			}

			$unitTestArray | Add-Member -NotePropertyName Details -NotePropertyValue $detailsArray

			$returnArray += $unitTestArray
		}

	    # trace Success
	    Write-LogDebug "Success Get-TfsUnitTestsInfo"
    }
    catch
    {
        $returnArray = $null

        # log Error 
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $returnArray
} 

#endregion
