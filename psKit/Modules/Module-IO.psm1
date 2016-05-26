<#
.Synopsis
Input & Output helper

.Description
contain IO fonctionnality : log, xml, IO, pskit..

.Notes
	fileName	: Module-IO.psm1
	version		: 0.31
	author		: Armand Lacore
#>

#region PARAMETERS

# define root folder, this module is meant to be hosted in Modules folder child of Root, Root\Modules\Module-IO.psm1
if (!$global:RootFolder) { $global:RootFolder = (Split-Path $PSScriptRoot) }

# define kit folders : Assemblies (.dll), Data (.xml), Modules (.psm1), Logs (.log)
if (!$global:AssembliesFolder) { $global:AssembliesFolder = [IO.Path]::Combine($global:RootFolder, "Assemblies") }
if (!$global:DataFolder) { $global:DataFolder = [IO.Path]::Combine($global:RootFolder, "Data") }
if (!$global:ModulesFolder) { $global:ModulesFolder = [IO.Path]::Combine($global:RootFolder, "Modules") }
if (!$global:LogsFolder) { $global:LogsFolder = [IO.Path]::Combine($global:RootFolder, "Logs") }

# create Logs folder if it does not exists
if (!(Test-Path $global:LogsFolder)) { New-Item $global:LogsFolder -ItemType Directory -Force | Out-Null }
		
# define log file if not set
if (!$global:LogFile) { $global:LogFile = [IO.Path]::Combine($global:LogsFolder, "$($env:USERNAME)_$(Get-Date -Format 'yyyy.MM.dd_HH.mm.ss').log") }

# variables used to get Write-LogError trace history
$global:TraceBuffer = [String]::Empty
$global:ErrorId = [String]::Empty

# Write-ObjectToHtml use HtmlEncode static method from System.Web.HttpUtility class
Add-Type -Assembly System.Web

#endregion

#region LOGGIN

<#
.Synopsis
add text to each line of a given text

.Description
split text into array with separator `n, then add text to each line

.Parameter Text
original text to append

.Parameter LineBegin
text to add to the begining of each line

.Parameter LineEnd
text to add to the end of each line

.Example
Add-ToEachLine $text "`t`t"

.Outputs
$Message
#>
Function Add-ToEachLine
{
	Param ( [Parameter(Mandatory=$true,Position=0)][AllowEmptyString()][string]$Text,
			[Parameter(Mandatory=$false,Position=1)][string]$LineBegin=[String]::Empty,
			[Parameter(Mandatory=$false,Position=2)][string]$LineEnd=[String]::Empty )

	#initialize return value
	$returnText = [String]::Empty
	
	try
	{
		$Text.Split("`n") | % { $returnText += "$($LineBegin)$($_)$($LineEnd)`n" }
	}
	catch
	{
	    # log Error
	    Write-LogError $($_.Exception) $($_.InvocationInfo)
	}

    return $returnText
}

<#
.Synopsis
format message

.Description
format message before they are logged

.Parameter Message
Content of the message to trace

.Parameter Title
Title to insert befor content message

.Parameter LineBefore
insert blank line before text

.Parameter LineAfter
insert blank line after text

.Parameter LineAfterTitle
insert blank line between title and message content

.Example
Format-Message "bonjour"

.Example
Format-Message "bonjour" "titre" -LineBefore

.Outputs
$Message
#>
Function Format-Message
{
	Param ( [Parameter(Mandatory=$true,Position=0)][AllowEmptyString()][string]$Message,
			[Parameter(Mandatory=$false,Position=1)][string]$Title,
			[Parameter(Mandatory=$false,Position=2)][switch]$LineBefore,
			[Parameter(Mandatory=$false,Position=3)][switch]$LineAfter,
			[Parameter(Mandatory=$false,Position=4)][switch]$LineAfterTitle )

	#initialize return value
	$toWrite = [String]::Empty
	
	try
	{
		if ($LineBefore)
		{
			$toWrite += [Environment]::NewLine
		}

		if ($Title)
		{
			$toWrite += "`t`t$Title$([Environment]::NewLine)"

			if ($LineAfterTitle)
			{
				$toWrite += [Environment]::NewLine
			}
		}

		$toWrite += $Message

		if ($LineAfter)
		{
			$toWrite += [Environment]::NewLine
		}
	}
	catch
	{
	    # log Error
	    Write-LogError $($_.Exception) $($_.InvocationInfo)
	}

    return $toWrite
}

<#
.Synopsis
enlarge UI
 
.Description
enlarge UI to get better console message formating

.Parameter Width
UI Width in pixel

.Parameter Height
UI Height in pixel

.Example
Set-UISize
#>
Function Set-UISize
{
	[CmdletBinding()]
	Param (	[Parameter(Mandatory=$false,Position=0)][int]$Width = 150,
			[Parameter(Mandatory=$false,Position=1)][int]$Height = 50 )

	Write-LogDebug "Start Set-UISize"
	
	try
	{
		# 3000 lines height is the default buffer height value
		$host.UI.RawUI.BufferSize = New-Object -TypeName System.Management.Automation.Host.Size -ArgumentList @($Width, 3000)
		$host.UI.RawUI.WindowSize = New-Object -TypeName System.Management.Automation.Host.Size -ArgumentList @($Width, $Height)
		
		# trace Success
		Write-LogDebug "Success Set-UISize"
	}
	catch
	{
		# log Error
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}
}

<#
.Synopsis
define host color

.Description
set fore and background color for all messages type

.Parameter ErrorColor
foreground color of error log event

.Parameter WarningColor
foreground color of warning log event

.Parameter DebugColor
foreground color of debug log event

.Parameter VerboseColor
foreground color of verbose log event

.Parameter ProgressColor
foreground color of progress event

.Example
Set-LogColor

.Example
Set-LogColor -ErrorColor "DarkCyan" -WarningColor "DarkRed" -DebugColor "DarkGreen" -VerboseColor "Gray" -ProgressColor "White"
#>
Function Set-LogColor
{
	# Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray, DarkGray, Blue, Green, Cyan, Red, Magenta, Yellow, White
	Param (	[Parameter(Mandatory=$false,Position=0)][ConsoleColor]$ErrorColor = "Red",
			[Parameter(Mandatory=$false,Position=1)][ConsoleColor]$WarningColor = "Yellow",
			[Parameter(Mandatory=$false,Position=2)][ConsoleColor]$DebugColor = "DarkGray",
			[Parameter(Mandatory=$false,Position=3)][ConsoleColor]$VerboseColor = "Gray",
			[Parameter(Mandatory=$false,Position=4)][ConsoleColor]$ProgressColor = "Green",
			[Parameter(Mandatory=$false,Position=5)][ConsoleColor]$BackgroundColor = "DarkMagenta" )

	try
	{
		$hostPrivateData = $host.PrivateData

		$hostPrivateData.ErrorForegroundColor = $ErrorColor
		$hostPrivateData.WarningForegroundColor = $WarningColor
		$hostPrivateData.DebugForegroundColor = $DebugColor
		$hostPrivateData.VerboseForegroundColor = $VerboseColor
		$hostPrivateData.ProgressForegroundColor = $ProgressColor

		$hostPrivateData.ErrorBackgroundColor = $BackgroundColor
		$hostPrivateData.WarningBackgroundColor = $BackgroundColor
		$hostPrivateData.DebugBackgroundColor = $BackgroundColor
		$hostPrivateData.VerboseBackgroundColor = $BackgroundColor
		$hostPrivateData.ProgressBackgroundColor = $BackgroundColor
	}
	catch
	{
		# log warning
		Write-LogWarning "Error Set-LogColor -Exception $($_.Exception.Message)"
	}
}

<#
.Synopsis
define log level

.Description
set error, warning, verbose, debug and trace preference to define log level

.Parameter $Level
Host, Error, Debug, Trace

.Example
Set-LogLevel "Host"
#>
Function Set-LogLevel
{
	Param (	[Parameter(Mandatory=$false,Position=0)][ValidateSet("Host","Verbose","Debug","Trace")][string]$Level )
	
	# Preference : Stop - display and stop, Inquire - display and prompt, Continue - display and continue, SilentlyContinue - do not display and continue, Suspend - suspend work (only for errors on workflows)
	# Trace : 0 - Turn script tracing off, 1 - Trace script lines as they are executed, 2 - Trace script lines, variable assignments, Function calls, and scripts

	try
	{
		switch ($Level)
		{
			"Host"
			{
				$global:ErrorActionPreference = "Stop"
				$global:WarningPreference = "Continue"
				$global:VerbosePreference = "SilentlyContinue"
				$global:DebugPreference = "SilentlyContinue"
				Set-PSDebug -Trace 0
			}
			"Verbose"
			{
				$global:ErrorActionPreference = "Stop"
				$global:WarningPreference = "Continue"
				$global:VerbosePreference = "Continue"
				$global:DebugPreference = "SilentlyContinue"
				Set-PSDebug -Trace 0
			}
			"Debug"
			{
				$global:ErrorActionPreference = "Stop"
				$global:WarningPreference = "Continue"
				$global:VerbosePreference = "Continue"
				$global:DebugPreference = "Continue"
				Set-PSDebug -Trace 0
			}
			"Trace"
			{
				$global:ErrorActionPreference = "Stop"
				$global:WarningPreference = "Continue"
				$global:VerbosePreference = "Continue"
				$global:DebugPreference = "Continue"
				Set-PSDebug -Trace 1
			}
		}
	}
	catch
	{
		# log warning
		Write-LogWarning "Error Set-LogLevel -Exception $($_.Exception.Message)"
	}
}

<#
.Synopsis
trace Debug

.Description
send trace to host as Debug stream and add content to logfile

.Parameter Message
Content of the message to trace

.Parameter Title
Title to insert befor content message

.Parameter LineBefore
insert blank line before text

.Parameter LineAfter
insert blank line after text

.Parameter LineAfterTitle
insert blank line between title and message content

.Example
Write-LogDebug -Message "bonjour"

.Outputs
$Message
#>
Function Write-LogDebug
{
	Param ( [Parameter(Mandatory=$true,Position=0)][AllowEmptyString()][string]$Message,
			[Parameter(Mandatory=$false,Position=1)][string]$Title,
			[Parameter(Mandatory=$false,Position=2)][switch]$LineBefore,
			[Parameter(Mandatory=$false,Position=3)][switch]$LineAfter,
			[Parameter(Mandatory=$false,Position=4)][switch]$LineAfterTitle )

	try
	{
		if ($global:DebugPreference -ne "SilentlyContinue")
		{
			$toWrite = Format-Message "$($Message) -Time $(Get-Date -Format 'HH.mm.ss.fff')" $Title $LineBefore $LineAfter $LineAfterTitle

			Add-Content $toWrite -Path $global:LogFile
			Write-Debug $toWrite
		}
	}
	catch
	{
		# log Error
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}
}

<#
.Synopsis
log Error

.Description
send log to host as Error stream and add content to logfile

.Parameter Exception
Exception object, used to get error message

.Parameter InvocationInfo
Invocation object info, use to get script name and line

.Parameter Title
Title to insert befor content message

.Parameter LineAfterTitle
insert blank line between title and message content

.Example
Write-LogError $($_.Exception) $($_.InvocationInfo)

.Outputs
$Message
#>
Function Write-LogError
{
	Param ( [Parameter(Mandatory=$true,Position=0)][System.Exception]$Exception,
            [Parameter(Mandatory=$true,Position=1)][System.Management.Automation.InvocationInfo]$InvocationInfo,
			[Parameter(Mandatory=$false,Position=2)][string]$Title,
			[Parameter(Mandatory=$false,Position=3)][switch]$LineAfterTitle )

	if ($global:ErrorActionPreference -ne "SilentlyContinue")
	{
		$source = Split-Path $InvocationInfo.ScriptName -leaf

		if ($Exception.Message.Length -gt 33 -and $Exception.Message.Substring(0,33) -eq $global:ErrorId)
		{
			$traceText = "$global:TraceBuffer -Line $($InvocationInfo.ScriptLineNumber)"

			$global:TraceBuffer = "-Trace $source"

			Add-Content $traceText -Path $global:LogFile
			Write-Error "$($Exception.Message)$([Environment]::NewLine)$($traceText)$([Environment]::NewLine)"
		}
		else
		{
			[string]$exceptionType = $Exception.GetType()
			[string]$text = "at $(Get-Date -Format 'HH.mm.ss.fff') an $exceptionType exception was thrown :$([Environment]::NewLine)"

			$message = ($InvocationInfo | Select-Object MyCommand, ScriptLineNumber, ScriptName, Line, OffsetInLine | Format-List | Out-String)
			$message += [Environment]::NewLine

			switch ($exceptionType)
			{
				"System.Reflection.ReflectionTypeLoadException" { $message += ($Exception | Select-Object Message, Source, LoaderExceptions, StackTrace | Format-List | Out-String) }
				default { $message += ($Exception | Select-Object Message, Source, ErrorCode, InnerException, StackTrace | Format-List | Out-String) }
			}

			$text += Format-Message $message $Title $false $false $LineAfterTitle

			$global:ErrorId = $text.Substring(0,33)
			$global:TraceBuffer = "-Trace $source"

			Add-Content $text -Path $global:LogFile
			Write-Error "$text"
		}
	}
}

<#
.Synopsis
trace event

.Description
send log message to windows events log

.Parameter Message
Content of the message to log

.Parameter Type
message type as string : Information, Error, Warning, SuccessAudit, FailureAudit

.Parameter Source
Event source (will be created if not exist)

.Parameter LogName
destination log : Application, Security, Setup, System

.Parameter EventID
id of the event

.Parameter Computer
destination computer

.Example
Write-LogEvent "bonjour"

.Example
Write-LogEvent "bonjour" "Warning" "Batch"

.Outputs
$Message
#>
Function Write-LogEvent
{
	[CmdletBinding()]
	Param ( [Parameter(Mandatory=$true,Position=0)][string]$Message,
			[Parameter(Mandatory=$false,Position=1)][string]$Type="Information",
			[Parameter(Mandatory=$false,Position=2)][string]$Source="myPSKit",
			[Parameter(Mandatory=$false,Position=3)][string]$LogName="IO",
			[Parameter(Mandatory=$false,Position=5)][int]$EventID=0,
			[Parameter(Mandatory=$false,Position=4)][string]$ComputerName="localhost" )

	try
	{
		Write-LogVerbose "updating source $Source on $LogName log, host : $($ComputerName)"

		# remove port from SqlServer computername
		if ($ComputerName -match ",") { $ComputerName = $ComputerName -replace '(.*),(.*)', '$1' }

		try
		{
			# check for eventLog
			Get-EventLog -LogName $LogName -Source $Source -ComputerName $ComputerName -ErrorAction Stop | Select-Object -First 1
		}
		catch
		{
			Write-LogVerbose "Creating source $Source on $LogName log, host : $($ComputerName)"

			# create log entry if it does not exists
			New-EventLog -LogName $LogName -Source $Source -ComputerName $ComputerName
		}

		Write-LogEventLog -LogName $LogName -Source $Source -ComputerName $ComputerName -EventId $EventID -Message $Message -EntryType $Type
	}
	catch
	{
		# log warning
		Write-LogWarning "Error Write-LogEvent -Exception $($_.Exception.Message)"
	}
}

<#
.Synopsis
trace Host message

.Description
send trace to host as Host stream and add content to logfile

.Parameter Message
Content of the message to trace

.Parameter Title
Title to insert befor content message

.Parameter LineBefore
insert blank line before text

.Parameter LineAfter
insert blank line after text

.Parameter LineAfterTitle
insert blank line between title and message content

.Example
Write-LogHost -Message "bonjour"

.Outputs
$Message
#>
Function Write-LogHost
{
	Param ( [Parameter(Mandatory=$true,Position=0)][AllowEmptyString()][string]$Message,
			[Parameter(Mandatory=$false,Position=1)][string]$Title,
			[Parameter(Mandatory=$false,Position=2)][switch]$LineBefore,
			[Parameter(Mandatory=$false,Position=3)][switch]$LineAfter,
			[Parameter(Mandatory=$false,Position=4)][switch]$LineAfterTitle )

	try
	{
		$toWrite = Format-Message $Message $Title $LineBefore $LineAfter $LineAfterTitle

		Add-Content $toWrite -Path $global:LogFile
		Write-Host $toWrite
	}
	catch
	{
		# log Error
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}
}

<#
.Synopsis
ObjectFormat object report

.Description
use Object Format-List, Trim() and $([Environment]::NewLine) to get a proper object report

.Parameter Object
object to log

.Parameter Message
title message for the object

.Parameter Level
log level output

.Parameter Format
object output formatting

.Parameter SkipLineBefore
do not insert blank line before object

.Parameter SkipLineAfter
do not insert blank line after object

.Parameter LineAfterTitle
insert blank line between title and message content

.Example
Write-LogObject $object

.Example
Write-LogObject -Object $object -Format "Table"

.Outputs
$Message
#>
Function Write-LogObject
{
	Param ( [Parameter(Mandatory=$true,Position=0)][AllowNull()][object]$Object,
			[Parameter(Mandatory=$false,Position=1)][string]$Title,
			[Parameter(Mandatory=$false,Position=2)][ValidateSet("Host","Verbose","Warning","Debug","Error")][string]$Level="Verbose",
			[Parameter(Mandatory=$false,Position=3)][ValidateSet("List","Table","Wide")][string]$Format="List",
			[Parameter(Mandatory=$false,Position=4)][switch]$SkipLineBefore,
			[Parameter(Mandatory=$false,Position=5)][switch]$SkipLineAfter,
			[Parameter(Mandatory=$false,Position=6)][switch]$LineAfterTitle )
	try
	{
		if ($Object)
		{
			switch ($Format)
			{
				"List"
				{ $toWrite = ($Object | Format-List | Out-String).Trim() }
				"Table"
				{ $toWrite = ($Object | Format-Table | Out-String).Trim() }
				"Wide"
				{ $toWrite = ($Object | Format-Wide | Out-String).Trim() }
			}
		}
		else
		{
		    $toWrite = "NULL OBJECT"
		}

		switch ($Level)
		{
			"Host"
			{ Write-LogHost $toWrite $Title (!$SkipLineBefore) (!$SkipLineAfter) $LineAfterTitle }
			"Verbose"
			{ Write-LogVerbose $toWrite $Title (!$SkipLineBefore) (!$SkipLineAfter) $LineAfterTitle }
			"Warning"
			{ Write-LogWarning $toWrite $Title (!$SkipLineBefore) (!$SkipLineAfter) $LineAfterTitle }
			"Debug"
			{ Write-LogDebug $toWrite $Title (!$SkipLineBefore) (!$SkipLineAfter) $LineAfterTitle }
			"Error"
			{ Throw $(Format-Message $toWrite $Title (!$SkipLineBefore) (!$SkipLineAfter) $LineAfterTitle) }
		}
	}
	catch
	{
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}
}

<#
.Synopsis
trace Verbose

.Description
send trace to host as Verbose stream and add content to logfile

.Parameter Message
Content of the message to trace

.Parameter Title
Title to insert befor content message

.Parameter LineBefore
insert blank line before text

.Parameter LineAfter
insert blank line after text

.Parameter LineAfterTitle
insert blank line between title and message content

.Example
Write-LogVerbose -Message "bonjour"

.Outputs
$Message
#>
Function Write-LogVerbose
{
	Param ( [Parameter(Mandatory=$true,Position=0)][AllowEmptyString()][string]$Message,
			[Parameter(Mandatory=$false,Position=1)][string]$Title,
			[Parameter(Mandatory=$false,Position=2)][switch]$LineBefore,
			[Parameter(Mandatory=$false,Position=3)][switch]$LineAfter,
			[Parameter(Mandatory=$false,Position=4)][switch]$LineAfterTitle )

	try
	{
		if ($global:VerbosePreference -ne "SilentlyContinue")
		{
			$toWrite = Format-Message $Message $Title $LineBefore $LineAfter $LineAfterTitle

			Add-Content $toWrite -Path $global:LogFile
			Write-Verbose $toWrite
		}
	}
	catch
	{
		# log Error
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}
}

<#
.Synopsis
log Warning

.Description
send log to host as Warning stream and add content to logfile

.Parameter Message
Content of the message to trace

.Parameter Title
Title to insert befor content message

.Parameter LineBefore
insert blank line before text

.Parameter LineAfter
insert blank line after text

.Parameter LineAfterTitle
insert blank line between title and message content

.Example
Write-LogWarning -Message "bonjour"

.Outputs
$Message
#>
Function Write-LogWarning
{
	Param ( [Parameter(Mandatory=$true,Position=0)][AllowEmptyString()][string]$Message,
			[Parameter(Mandatory=$false,Position=1)][string]$Title,
			[Parameter(Mandatory=$false,Position=2)][switch]$LineBefore,
			[Parameter(Mandatory=$false,Position=3)][switch]$LineAfter,
			[Parameter(Mandatory=$false,Position=4)][switch]$LineAfterTitle )

	try
	{
		if ($global:WarningPreference -ne "SilentlyContinue")
		{
			$toWrite = Format-Message $Message $Title $LineBefore $LineAfter $LineAfterTitle

			Add-Content $toWrite -Path $global:LogFile
			Write-Warning $toWrite
		}
	}
	catch
	{
		# log Error
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}
}

<#
.Synopsis
Write a Objects array to html table

.Description
Write a one level Objects array into html table

.Parameter Object
array containing Objects to report

.Parameter TableClass
if set, define the table class

.Parameter ThClass
if set, define the table header class

.Parameter TrClass
if set, define the table row class

.Parameter TdClass
if set, define the table data class

.Parameter Horizontal
switch, set to true if you want to print html table horizontaly

.Parameter SkipTh
switch, set to true if you don't want to print table header

.Example
Write-ObjectToHtml $object -FirstLineAsTh -ThClass "bigTitle"

.Outputs
table on success / null on fail
#>
Function Write-ObjectToHtml
{
	[CmdletBinding()]
	Param (	[Parameter(Mandatory=$true,Position=0)][AllowNull()][object]$Object,
			[Parameter(Mandatory=$false,Position=1)][string]$TableClass,
			[Parameter(Mandatory=$false,Position=2)][string]$ThClass,
			[Parameter(Mandatory=$false,Position=3)][string]$TrClass,
			[Parameter(Mandatory=$false,Position=4)][string]$TdClass,
			[Parameter(Mandatory=$false,Position=5)][string]$TableId,
            [Parameter(Mandatory=$false,Position=6)][switch]$Horizontal,
            [Parameter(Mandatory=$false,Position=7)][switch]$SkipTh,
            [Parameter(Mandatory=$false,Position=8)][switch]$EncodeHtml	)

	Write-LogDebug "Start Write-ObjectToHtml"
	
	try
	{
		Write-LogVerbose "Saving custom PSObjects to HTML table"
        
        if ($Object)
        {
            # initialize return value
		    $htmlTable = [String]::Empty
        
            if ($TableClass) { $TableClass = " class=`"$TableClass`"" } else { $TableClass = [String]::Empty }
            if ($ThClass) { $ThClass = " class=`"$ThClass`"" } else { $ThClass = [String]::Empty }
            if ($TrClass) { $TrClass = " class=`"$TrClass`"" } else { $TrClass = [String]::Empty }
            if ($TdClass) { $TdClass = " class=`"$TdClass`"" } else { $TdClass = [String]::Empty }
            if ($TableId) { $TableId = " Id=`"$TableId`"" } else { $TableId = [String]::Empty }

		    if ($Object -is [array])
		    {
			    $testLine = $Object | Select -First 1
		    }
		    else
		    {
			    $testLine = $Object
		    }

		    if ($testLine -is [Collections.Hashtable] -or $testLine -is [Collections.Specialized.OrderedDictionary])
		    {
			    $propertiesName = $testLine.Keys
		    }
		    elseif ($testLine -is [PSCustomObject])
		    {
			    $propertiesName = $testLine | Get-Member -MemberType NoteProperty
		    }
		    else
		    {
			    Throw "this method only allow PsCustomObject or Hashtable content"
		    }
            
		    # write table start
            $htmlTable += "<table$($TableId)$($TableClass)>$([Environment]::NewLine)"
			
            # write table header
		    if (!$SkipTh -and !$Horizontal)
		    {
                $htmlTable += "`t<tr$($TrClass)>"

			    foreach ($propertyName in $propertiesName)
			    {
                    $htmlTable += "<th$($ThClass)>$propertyName</th>"
			    }

			    $htmlTable += "</tr>$([Environment]::NewLine)"
		    }

		    # write table rows
		    foreach ($entity in $Object)
		    {
                if (!$Horizontal)
                {
                    $htmlTable += "`t<tr$($TrClass)>"

			        foreach ($propertyName in $propertiesName)
			        {
                        if ($EncodeHtml) { $propertyValue = $([System.Web.HttpUtility]::HtmlEncode($entity.$propertyName)) } else { $propertyValue = $entity.$propertyName }

                        $htmlTable += "<td$($TdClass)>$propertyValue</td>"
			        }

			        $htmlTable += "</tr>$([Environment]::NewLine)"
                }
                else
                {
			        foreach ($propertyName in $propertiesName)
			        {
                        if ($EncodeHtml) { $propertyValue = $([System.Web.HttpUtility]::HtmlEncode($entity.$propertyName)) } else { $propertyValue = $entity.$propertyName }

                        $htmlTable += "`t<tr$($TrClass)>"

                        if (!$SkipTh)
                        {
                            $htmlTable += "<th$($TdClass)>$propertyName</th><td$($TdClass)>$propertyValue</td>"
                        }
                        else
                        {
                            $htmlTable += "<td$($TdClass)>$propertyValue</td>"
                        }

                        $htmlTable += "</tr>$([Environment]::NewLine)" 
                    }
                }
		    }

		    $htmlTable += "</table>"
        }
        else
        {
            Write-LogWarning "Object is null"
            $htmlTable = [String]::Empty
        }

		# trace Success
		Write-LogDebug "Success Write-ObjectToHtml"
	}
	catch
	{
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
	}

	return $htmlTable
}

<#
.Synopsis
Write a psObject to xml elements

.Description
Write a one level psObject into xml elements

.Parameter Object
psObject to report

.Parameter Path
xml output path

.Parameter RootName
xml root name

.Parameter ElementName
xml element name 

.Example
Write-ObjectToXml -Object $psObject -Path "C:\output.xml"

.Outputs
true on success / false on fail
#>
Function Write-ObjectToXml
{
	[CmdletBinding()]
	Param (	[Parameter(Mandatory=$true,Position=0)][PSObject]$Object,
			[Parameter(Mandatory=$true,Position=1)][string]$Path,
			[Parameter(Mandatory=$false,Position=2)][string]$Root="Document" )

	Write-LogDebug "Start Write-ObjectToXml"
	
	try
	{
		Write-LogVerbose "Saving custom PSObject to XML file $($Path)"

		$folder = (Split-Path $Path)

		if (!(Test-Path $folder))
		{
			New-Item -Path $folder -ItemType Directory -Force
		}

		$xmlWriter = New-Object -TypeName System.Xml.XmlTextWriter -ArgumentList @($Path, [Text.Encoding]::GetEncoding("UTF-8"))
		$xmlWriter.Formatting = "Indented"
		$xmlWriter.Indentation = 1
		$xmlWriter.IndentChar = "`t"
		
		$xmlWriter.WriteStartDocument()
		
		$xmlWriter.WriteStartElement($Root)

		$members = $Object | Get-Member -MemberType NoteProperty

		foreach ($member in $members)
		{
			$name = $member.Name
			$value = $member.Definition -replace "^(.*)$name=(.*)$", '$2'

			$xmlWriter.WriteElementString($name, $value)
		}

		$xmlWriter.WriteEndElement()
		$xmlWriter.WriteEndDocument()
		$xmlWriter.Flush()
		$xmlWriter.Close()

		# trace Success
		Write-LogDebug "Success Write-ObjectToXml"
	}
	catch
	{
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
		$xml = $null
	}

	return $xml
}

#endregion

#region XML

<#
.Synopsis
Get Xml element or attribute value

.Description
Get an xml element or attribute identified by XPath

.Parameter File
full path to the file edited

.Parameter XPath
Xpath to property to get

.Example
Get-XPathValue "C:\example.xml" "/system.web/authentication/@mode"

.Outputs
value on success / null on fail
#>
Function Get-XPathValue
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][string]$File,
			[Parameter(Mandatory=$true,Position=1)][string]$XPath )

    Write-LogDebug "Start Get-XPathValue"

    #initialize return value
	$returnValue = $null

    try
    {
		Write-LogVerbose "opening $File to get $XPath"

		# get xml file
		$xmlFile = [xml](Get-Content $File)

		# get node corresponding to XPath
		$xmlNode = $xmlFile.SelectSingleNode($XPath)
        
        if ($xmlNode)
        {
            if ($xmlNode -is [Xml.XmlAttribute])
            {
                $returnValue = $xmlNode.Value
            }
            elseif ($xmlNode -is [Xml.XmlElement])
            {
                $returnValue = $xmlNode.InnerText
            }
        }

	    # trace Success
	    Write-LogDebug "Success Get-XPathValue"
    }
    catch
    {
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $returnValue
}

<#
.Synopsis
Set Xml element or attribute value

.Description
Set or (un)comment an xml element or attribute identified by XPath

.Parameter File
full path to the file edited

.Parameter XPath
Xpath to property to edit

.Parameter Value
Value to set

.Parameter Type
Edition type : setValue, add, remove, replace, comment, unComment

.Parameter addMode
insertion mode : before \ after Xpath given

.Example
Set-XPathValue "C:\example.xml" "/system.web/authentication/@mode" "true"

.Example
Set-XPathValue "C:\example.xml" "/system.web/authentication/@mode" -Type "comment"

.Outputs
true on success / false on fail
#>
Function Set-XPathValue
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][string]$File,
			[Parameter(Mandatory=$true,Position=1)][string]$XPath,
			[Parameter(Mandatory=$false,Position=2)][AllowNull()][string]$Value,
			[Parameter(Mandatory=$false,Position=3)][AllowNull()][ValidateSet("","setValue", "add", "remove", "replace", "comment", "unComment")][string]$Type="setValue",
			[Parameter(Mandatory=$false,Position=4)][AllowNull()][ValidateSet("","before", "after", "in")][string]$AddMode="in")

    Write-LogDebug "Start Set-XPathValue"

    #initialize return value
    [bool]$returnValue = $true

    try
    {
        switch ($Type)
        {
            "setValue"
            {
                $message = "to set Value $Value on $XPath"
            }
            "add"
            {
                $message = "to add $Value $AddMode on $XPath"
            }
            "remove"
            {
                $message = "to remove $XPath"
            }
            "replace"
            {
                $message = "to replace $XPath with $Value"
            }
            "comment"
            {
                $message = "to comment $XPath"
            }
            "unComment"
            {
                $message = "to uncomment $XPath"
            }
        }

        Write-LogVerbose "editing file $File $message"

		# get xml file
		$xmlFile = [xml](Get-Content $File)

		# get node corresponding to XPath
		if ($Type -ne "unComment")
		{
			$xmlNode = $xmlFile.SelectSingleNode($XPath)
		}
		else
		{
			$parentXPath = Get-InString $XPath '/' -Last
			$childXPath = $XPath -replace "$parentXPath/", ''

			$xmlNode = $xmlFile.SelectNodes("$parentXPath/comment()") | ? { $_.InnerText -match "<$childXPath" }
		}

		# test node availability
		if (!$xmlNode)
		{
			Throw "Element referenced under $XPath does not exist"
		}

		# process node operations
		switch ($Type)
		{
			"add"
			{
				Write-LogVerbose "add $Value to $XPath"

				$tempNode = $xmlFile.CreateElement("tempNode")
				$tempNode.InnerXml = $Value

                if ($AddMode -eq "in")
                {
                    $xmlNode.AppendChild($tempNode.FirstChild)
                }
				elseif ($AddMode -eq "after")
				{
					$xmlNode.ParentNode.InsertAfter($tempNode.FirstChild, $xmlNode)
				}
				elseif ($AddMode -eq "before")
				{
					$xmlNode.ParentNode.InsertBefore($tempNode.FirstChild, $xmlNode)
				}
				else
				{
					Throw "add mode does not exist"
				}
			}
			"remove"
			{
				Write-LogVerbose "remove $XPath"

				$xmlNode.ParentNode.RemoveChild($xmlNode)
			}
			"replace"
			{
				Write-LogVerbose "replace $XPath with $Value"

				$tempNode = $xmlFile.CreateElement("tempNode")
				$tempNode.InnerXml = $Value

				$xmlNode.ParentNode.ReplaceChild($tempNode.FirstChild, $xmlNode)
			}
			"setValue"
			{
				Write-LogVerbose "setting $XPath to $Value"
                
                if ($xmlNode -is [Xml.XmlAttribute])
                {
                    $xmlNode.Value = $Value
                }
                elseif ($xmlNode -is [Xml.XmlElement])
                {
                    $xmlNode.InnerText = $value
                }
			}
			"comment"
			{
				Write-LogVerbose "comment $XPath"

				$xmlComment = $xmlFile.CreateComment($xmlNode.OuterXml)
				$xmlNode.ParentNode.ReplaceChild($xmlComment, $xmlNode)
			}
			"unComment"
			{
				Write-LogVerbose "unComment $XPath"

				$tempNode = $xmlFile.CreateElement("tempNode")
				$tempNode.InnerXml = $xmlNode.Value

				$xmlNode.ParentNode.ReplaceChild($($tempNode.SelectSingleNode("/$childXPath")), $xmlNode)
			}
		}

		# save file
		$xmlFile.Save($File)

	    # trace Success
	    Write-LogDebug "Success Set-XPathValue"
    }
    catch
    {
        $returnValue = $false

        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $returnValue
}

<#
.Synopsis
Test Xpath existence

.Description
send true if XPath is found on File, false if not

.Parameter File
full path to the file edited

.Parameter XPath
Xpath to test

.Example
Test-XPath "C:\example.xml" "/system.web/authentication/@mode"

.Outputs
true if XPath exist / false if XPath does not exist
#>
Function Test-XPath
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][string]$File,
			[Parameter(Mandatory=$true,Position=1)][string]$XPath )

    Write-LogDebug "Start Test-XPath"

    #initialize return value
	$returnValue = $false

    try
    {
		Write-LogVerbose "opening $File to test if $XPath exists"

		# get xml file
		$xmlFile = [xml](Get-Content $File)

		# get node corresponding to XPath
		$xmlNode = $xmlFile.SelectSingleNode($XPath)
        
        if ($xmlNode)
        {
            $returnValue = $true
        }

	    # trace Success
	    Write-LogDebug "Success Test-XPath"
    }
    catch
    {
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $returnValue
}

#endregion

#region TOOLS

<#
.Synopsis
Create folders if they don't not exists

.Description
Create folders if they don't exists, using [IO.Directory]::Exists and New-Item -Force

.Parameter Folders
full path list of the folder to assert

.Example
Assert-Folders "H:\Logs\QA2"

.Example
Assert-Folders "H:\Logs\QA2", "F:\Data\QA2"
#>
Function Assert-Folders
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][array]$Folders )

    Write-LogDebug "Start Assert-Folders"

    try
    {
        foreach ($folder in $Folders)
        {
			if (![IO.Directory]::Exists($folder))
			{
				New-Item -Path $folder -ItemType Directory -Force -ErrorAction SilentlyContinue
			}
        }

	    # trace Success
	    Write-LogDebug "Success Assert-Folders"
    }
    catch
    {
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }
}

<#
.Synopsis
string extract fonction

.Description
Get string part delimited by a separator

.Parameter InString
String to seek

.Parameter Delimiter
characters delimiter

.Parameter SecondDelimiter
characters second delimiter

.Parameter Last
switch, if true get last occurence of delimiter (default = first)

.Parameter SecondLast
switch, if true get last occurence of second delimiter (default = first)

.Parameter Right
switch, if true get right part of the string (default = left, switch disabled if 2 delimiters)

.Parameter Up
switch, if true get uppercase result

.Parameter Low
switch, if true get lowercase result

.Output
searched string or blank if not found
#>
Function Get-InString
{
	[CmdletBinding()]
	Param ( [Parameter(Mandatory=$true,Position=0)][string]$InString,
			[Parameter(Mandatory=$true,Position=1)][string]$Delimiter,
			[Parameter(Mandatory=$false,Position=2)][string]$SecondDelimiter,
			[Parameter(Mandatory=$false,Position=3)][switch]$Last,
			[Parameter(Mandatory=$false,Position=4)][switch]$SecondLast,
			[Parameter(Mandatory=$false,Position=5)][switch]$Right,
			[Parameter(Mandatory=$false,Position=6)][switch]$Up,
			[Parameter(Mandatory=$false,Position=7)][switch]$Low )

	Write-LogDebug "Start Get-InString"

	# initialize return value
	$returnValue = [String]::Empty

	try
	{
		if ($SecondDelimiter)
		{
			if ($InString -match $Delimiter)
			{
				if ($Last)
				{
					$index = $InString.LastIndexOf($Delimiter)
				}
				else
				{
					$index = $InString.IndexOf($Delimiter)
				}

				$tmpValue = $InString.substring($index + 1, ($InString.Length - $index - 1))
			}
			else
			{
				$tmpValue = $InString
			}

			if ($tmpValue -match $SecondDelimiter)
			{
				if ($SecondLast)
				{
					$index = $tmpValue.LastIndexOf($SecondDelimiter)
				}
				else
				{
					$index = $tmpValue.IndexOf($SecondDelimiter)
				}

				$returnValue = $tmpValue.substring(0, $index)
			}
			else
			{
				$returnValue = $tmpValue
			}
		}
		else
		{
			if ($InString -match $Delimiter)
			{
				if ($Last)
				{
					$index = $InString.LastIndexOf($Delimiter)
				}
				else
				{
					$index = $InString.IndexOf($Delimiter)
				}

				if ($Right)
				{
					$returnValue = $InString.substring($index + 1, ($InString.Length - $index - 1))
				}
				else
				{
					$returnValue = $InString.substring(0, $index)
				}
			}
			else
			{
				$returnValue = $InString
			}
		}

		if ($Up)
		{
			$returnValue = $returnValue.ToUpper()
		}

		if ($Low)
		{
			$returnValue = $returnValue.ToLower()
		}

		# trace Success
		Write-LogDebug "Success Get-InString"
	}
	catch
	{
		# log Error
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}

	return $returnValue
}

<#
.Synopsis
prompt for user confirmation

.Description
send prompt for a Yes\No question

.Parameter Title
text for title

.Parameter Message
text for message

.Parameter YesCaption
text for yes option

.Parameter NoCaption
text for no caption

.Example
Get-Prompt "proceed installation ?"

.Outputs
$true or $false
#>
Function Get-Prompt
{
	[CmdletBinding()]
	Param (	[Parameter(Mandatory=$true,Position=0)][AllowEmptyString()][string]$Message="Do you wan't to continue ?",
			[Parameter(Mandatory=$false,Position=1)][string]$Title="Prompt user",
			[Parameter(Mandatory=$false,Position=2)][string]$YesCaption="Proceed",
			[Parameter(Mandatory=$false,Position=3)][string]$NoCaption="Cancel" )

	Write-LogDebug "Start Get-Prompt"
	
	try
	{
		$yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList @("&Yes", $YesCaption)
		$no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList @("&No", $NoCaption)
		$options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)

		$confirmationChoice = $host.ui.PromptForChoice($Title, $Message, $options, 0)

		if ($confirmationChoice -eq 0)
		{
			return $true
		}
		else
		{
			return $false
		}
	}
	catch
	{
		# log Error
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}
}

<#
.Synopsis
convert secure string

.Description
convert secure string to plain text string using Runtime.InteropServices.Marshal library

.Parameter SecureString
secure string to convert

.Example
Get-SecureString -SecureString $secureStr

.Outputs
Secure string as string
#>
Function Get-SecureString
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][Security.SecureString]$SecureString )

    Write-LogDebug "Start Get-SecureString"

    try
    {
		return [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString))

	    # trace Success
	    Write-LogDebug "Success Get-SecureString"
    }
    catch
    {
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }
}

<#
.Synopsis
get sha256 hash

.Description
perform a sha256 compute then convert to String and return

.Parameter File
file to compute

.Output
content-md5 as base64 String \ null if error
#>
Function Get-SHA256
{
	[CmdletBinding()]
	Param ( [Parameter(Mandatory=$true,Position=0)][string]$File )

	Write-LogDebug "Start Get-SHA256"

	# initialize return value
	$fileHash = $null

	try
	{
        Write-LogVerbose "getting $(Split-Path $File -leaf) hash (sha256)"

		if (Test-Path $File)
		{
			$fileHash = $(Get-FileHash $File).Hash

			# trace Success
			Write-LogDebug "Success Get-SHA256"
		}
		else
		{
			# log warning
			Write-LogWarning "file : $($File) does not exist"
		}
	}
	catch
	{
		# log Error
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}

	return $fileHash
}

<#
.Synopsis
Create or clean temp folder

.Example
Get-TempFolder

.Outputs
$tempFolder on success, $null on fail
#>
Function Get-TempFolder
{
	Write-LogDebug "Start Get-TempFolder"
	
	try
	{
		$tempFolder = [IO.Path]::Combine($global:RootFolder, "Temp$(Get-Date -Format yyyyMMddHHmmssfff)")
        
        Write-LogVerbose "asserting temp folder : $tempFolder"

		# clean or create temp folder
		if (!(Test-Path $tempFolder)) { New-Item $tempFolder -ItemType Directory -Force | Out-Null }
		else {	Get-ChildItem $tempFolder | % { Remove-Item $_.FullName -Force -ErrorAction Continue } }
		
		Write-LogDebug "Success Get-TempFolder"
	}
	catch
	{
		$tempFolder = $null
		
		# log Error
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}
	
	return $tempFolder
}

<#
.Synopsis
remove folder

.Description
deals with remove-item bug with recurse and not empty folders

.Parameter FolderPath
full folder path

.Example
Remove-Folder "C:\Test"

.Outputs
none
#>
Function Remove-Folder
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][string]$FolderPath )

    Write-LogDebug "Start Remove-Folder"

	Write-LogVerbose "Removing folder : $FolderPath" -LineBefore
	
	if (Test-Path $FolderPath)
	{
		try
		{
			Remove-Item $FolderPath -Recurse -Force
		}
		catch
		{
			# remove item method often fail on not empty folder
			Get-ChildItem -Path $FolderPath -Recurse | Remove-Item -Force

			Remove-Item $FolderPath -Recurse -Force -ErrorAction Continue
		}
	}

    # trace Success
    Write-LogDebug "Success Remove-Folder"
}

<#
.Synopsis
convert string to secure string

.Description
convert string to secure string using ConvertTo-String

.Parameter String
string to convert

.Example
Set-SecureString "bonjour"

.Outputs
string as secure string
#>
Function Set-SecureString
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][string]$String )

    Write-LogDebug "Start Set-SecureString"

    try
    {
		return (ConvertTo-SecureString $String -AsPlainText -Force)

	    # trace Success
	    Write-LogDebug "Success Set-SecureString"
    }
    catch
    {
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }
}

#endregion

#region psKitUtilities

<#
.Synopsis
import modules with an easy syntax

.Description
Import modules by module short names array, exemple : Root\Modules\Module-Exemple.psm1 and Root\Modules\Module-Exemple2.psm1 can be imported with syntax : Add-Modules "Exemple","Exemple2"

.Parameter Names
Short names list of the modules to import (not mandatory, default = all modules in Root\Modules folder)

.Parameter Assert
switch : select if wan't to test module availabily before import

.Example
Add-Modules "Infra", "Install"
#>
Function Add-Modules
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$false,Position=0)][array]$Names,
			[Parameter(Mandatory=$false,Position=1)][switch]$Assert )

    Write-LogDebug "Start Add-Modules"

    try
    {
        if ($Names)
        {
            foreach ($name in $Names)
            {
				if (!$Assert -or !(Get-Module -Name "Module-$name"))
				{
					$modulePath = [IO.Path]::Combine($global:ModulesFolder, "Module-$name.psm1")
					Import-Module $modulePath -Force -Global
				}
            }
        }
        else
        {
			if ($Assert)
			{
				Get-ChildItem ($global:ModulesFolder) | ? {!(Get-Module -Name $_.BaseName)} | % { Import-Module -Name $_.FullName -Force -Global }
			}
			else
			{
				Get-ChildItem ($global:ModulesFolder) | % { Import-Module -Name $_.FullName -Force -Global }
			}
        }

	    # trace Success
	    Write-LogDebug "Success Add-Modules"
    }
    catch
    {
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }
}

<#
.Synopsis
import assemblies with an easy syntax

.Description
Import assemblies by assembly short name array, exemple : Root\Assemblies\Microsoft.ServiceBus.dll and Root\Modules\Microsoft.WindowsAzure.Configuration.dll can be imported with syntax : Add-Modules "ServiceBus","WindowsAzure.Configuration"

.Parameter Names
Short names list of the assemblies to import

.Example
Add-Assemblies "SqlServer.SqlEnum", "SqlServer.Smo"
#>
Function Add-Assemblies
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][array]$Names )

    Write-LogDebug "Start Add-Assemblies"

    try
    {
        foreach ($name in $Names)
        {
            $assemblyName = "Microsoft.$name.dll"
            $assemblyPath = [IO.Path]::Combine($global:AssembliesFolder, $assemblyName)

            Add-Type -Path $assemblyPath
        }

	    # trace Success
	    Write-LogDebug "Success Add-Assemblies"
    }
    catch
    {
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }
}

<#
.Synopsis
return xml data

.Description
test-path, then read and return xml file as object

.Parameter File
file name without extension

.Parameter Folder
file folder, not mandatory, default = Data folder

.Parameter IsFullPath
-switch set to true if $File contain file full path

.Example
Get-Data "Builds"
#>
Function Get-Data
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][string]$File,
			[Parameter(Mandatory=$false,Position=1)][string]$Folder=$global:DataFolder,
			[Parameter(Mandatory=$false,Position=2)][switch]$IsFullPath )

    Write-LogDebug "Start Get-Data"

    try
    {
        Write-LogVerbose "reading $File data file"

        if ($IsFullPath)
        {
            $filePath = $File
        }
        else
        {
            $filePath = [IO.Path]::Combine($Folder, "$File.xml")
        }

		try
		{
			$returnObject = [xml](Get-Content $filePath -Encoding UTF8)
		}
		catch [System.Management.Automation.ItemNotFoundException]
		{
			if ($File -eq "History") 
			{ New-HistoryXml | Out-Null; $returnObject = [xml](Get-Content $filePath) }
			elseif ($File -eq "Events") 
			{ New-EventsXml | Out-Null; $returnObject = [xml](Get-Content $filePath) }
			elseif ($File -eq "Versions") 
			{ New-VersionsXml | Out-Null; $returnObject = [xml](Get-Content $filePath) }
			else 
			{ Write-LogError $($_.Exception) $($_.InvocationInfo) }
		}

        return $returnObject

	    # trace Success
	    Write-LogDebug "Success Get-Data"
    }
    catch
    {
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }
}

#endregion
