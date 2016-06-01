<# 
.Synopsis
query build timed example

.Description
timed examples on QueryBuilds methods from definition specs, see https://msdn.microsoft.com/en-us/library/cc340540(v=vs.120).aspx & https://msdn.microsoft.com/en-us/library/microsoft.teamfoundation.build.client.ibuilddefinitionspec(v=vs.120).aspx

.Notes 
	fileName	: Example-TfsBuilds.psm1
	version		: 0.006
	author		: Armand Lacore
#>

# script parameters
$testOccurency = 2

# tfs parameters
$tfsCollectionUri = "https://tfs.axacolor.igo6.com/tfs/DefaultCollection"
$tfsProjectName = "COLOR-SP"
$tfsBuildName = "CI-IGO6-R1.2"
$tfsBuildNumber = "CI-IGO6-R1.2_20160531.1"

# import IO and TFS modules (Hosted in Root\Modules)
$modulesShortName = @("IO", "Tfs")
$modulesShortName | % { if (!(Get-Module -Name "Module-$_")) { Import-Module ([IO.Path]::Combine($PSScriptRoot, "Modules", "Module-$_.psm1")) -Force -Global } }

$buildServer = (Get-TfsConnection $tfsCollectionUri).BuildServer

Function TimeTfsQuery([AllowNull()]$BuildNumber, [AllowNull()]$InformationTypes, $QueryOptions)
{
	Function MakeItReadable($Property)
	{
		switch ($Property)
		{
			"*"
			{
				return "all"
			}
			$null
			{
				return "null"
			}
			default
			{
				return $Property
			}
		}
	}
	
	Write-LogHost "testing $tfsBuildNumber with Build Number $(MakeItReadable($BuildNumber)), InformationTypes $(MakeItReadable($InformationTypes)) & QueryOptions $QueryOptions"
	
	$stopwatch = New-Object -TypeName System.Diagnostics.Stopwatch
	$stopwatch.Start()

	$buildSpecification = $buildServer.CreateBuildDetailSpec($tfsProjectName, $tfsBuildName)
	
	$buildSpecification.BuildNumber = $BuildNumber
	$buildSpecification.InformationTypes = $InformationTypes
	$buildSpecification.QueryOptions = $QueryOptions

	$tfsBuilds = $buildServer.QueryBuilds($buildSpecification)
	$stopwatch.Stop()
	
	Write-LogHost "time : $($stopwatch.ElapsedMilliseconds) ms"
	$tfsBuilds.Builds | ConvertTo-Json | Add-Content ([IO.Path]::Combine($global:LogsFolder, "BN_$(MakeItReadable($BuildNumber))-IT_$(MakeItReadable($InformationTypes))-QO_$($buildSpecification.QueryOptions)_builds-$($stopwatch.ElapsedMilliseconds)_ms.json"))
}

Function SimpleCall
{
	$stopwatch = New-Object -TypeName System.Diagnostics.Stopwatch
	$stopwatch.Start()

	$tfsBuilds = $buildServer.QueryBuilds($tfsProjectName, $tfsBuildName)
	$stopwatch.Stop()
	
	Write-LogHost "time : $($stopwatch.ElapsedMilliseconds) ms"
	$tfsBuilds.Builds | ConvertTo-Json | Add-Content ([IO.Path]::Combine($global:LogsFolder, "SimpleOverload_builds-$($stopwatch.ElapsedMilliseconds)_ms.json"))
}

for ($i=0; $i -le $testOccurency; $i++)
{
	# TimeTfsQuery "*" $null "None"
	# TimeTfsQuery "*" $null "All"
	SimpleCall
	TimeTfsQuery "*" "*" "All"
	# TimeTfsQuery "*" "*" "None"
	# TimeTfsQuery $tfsBuildNumber $null "None"
	# TimeTfsQuery $tfsBuildNumber $null "All"
	# TimeTfsQuery $tfsBuildNumber "*" "None"
	# TimeTfsQuery $tfsBuildNumber "*" "All"
}