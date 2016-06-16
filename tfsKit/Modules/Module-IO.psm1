<#
.Synopsis
Input & Output helper

.Description
contain IO fonctionnality : log, xml, IO.

.Notes
	fileName	: Module-IO.psm1
	version		: 0.52
	author		: Armand Lacore
#>

#region PARAMETERS

# define root folder, this module is meant to be hosted in Root\Modules
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
$global:TraceBuffer = [string]::Empty
$global:ErrorId = [string]::Empty

# Write-ObjectToHtml use HtmlEncode static method from System.Web.HttpUtility class
Add-Type -Assembly System.Web

#endregion

#region LOGGIN

<#
.Synopsis
Add text to each line of a given text

.Description
Split text into array with separator `n, then add text before and/or after each line

.Parameter Text
Original text to append

.Parameter LineBegin
Text to add to the begining of each line

.Parameter LineEnd
Text to add to the end of each line

.Example
Add-ToEachLine $text "`t`t"

.Outputs
Original text appended with extra text
#>
Function Add-ToEachLine
{
	Param ( [Parameter(Mandatory=$true,Position=0)][AllowEmptyString()][string]$Text,
			[Parameter(Mandatory=$false,Position=1)][string]$LineBegin=[string]::Empty,
			[Parameter(Mandatory=$false,Position=2)][string]$LineEnd=[string]::Empty )

	#initialize return value
	$returnText = [string]::Empty
	
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
Format message

.Description
Format message before they are logged, adding Title and/or extra blank line

.Parameter Message
Content of the message to trace

.Parameter Title
Title to insert before content message

.Parameter LineBefore
Insert blank line before text

.Parameter LineAfter
Insert blank line after text

.Parameter LineAfterTitle
Insert blank line between title and message content

.Example
Format-Message "hello"

.Example
Format-Message "hello" "title" -LineBefore

.Outputs
Original text formatted
#>
Function Format-Message
{
	Param ( [Parameter(Mandatory=$true,Position=0)][AllowEmptyString()][string]$Message,
			[Parameter(Mandatory=$false,Position=1)][string]$Title,
			[Parameter(Mandatory=$false,Position=2)][switch]$LineBefore,
			[Parameter(Mandatory=$false,Position=3)][switch]$LineAfter,
			[Parameter(Mandatory=$false,Position=4)][switch]$LineAfterTitle )

	#initialize return value
	$toWrite = [string]::Empty
	
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
Set powershell UI size
 
.Description
Set powershell UI size and buffer to get better console message display

.Parameter Width
UI Width in pixel

.Parameter Height
UI Height in pixel

.Example
Set-UISize

.Example
Set-UISize 200 70
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
Set powershell UI title
 
.Description
Set UI Title from $host.UI object

.Parameter Title
Title to set

.Example
Set-UITitle

.Example
Set-UITitle "my title"
#>
Function Set-UITitle
{
	[CmdletBinding()]
	Param (	[Parameter(Mandatory=$true,Position=0)][string]$Title )

	Write-LogDebug "Start Set-UITitle"
	
	try
	{
		$host.UI.RawUI.WindowTitle = $Title
		
		# trace Success
		Write-LogDebug "Success Set-UITitle"
	}
	catch
	{
		# log Error
		Write-LogError $($_.Exception) $($_.InvocationInfo)
	}
}

<#
.Synopsis
Define host color

.Description
Set fore and background color for all messages type

.Parameter ErrorColor
Foreground color of error log message

.Parameter WarningColor
Foreground color of warning log message

.Parameter DebugColor
Foreground color of debug log message

.Parameter VerboseColor
Foreground color of verbose log message

.Parameter ProgressColor
Foreground color of progress message

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
Define log level

.Description
Set error, warning, verbose, debug and trace preference to define log level

.Parameter Level
Log level from : Host, Error, Debug, Trace

.Example
Set-LogLevel "Verbose"
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
Trace debug message

.Description
Send trace to host as Debug message and add content to logfile

.Parameter Message
Content of the message to trace

.Parameter Title
Title to insert before content message

.Parameter LineBefore
Insert blank line before text

.Parameter LineAfter
Insert blank line after text

.Parameter LineAfterTitle
Insert blank line between title and message content

.Example
Write-LogDebug "bonjour"

.Outputs
Print message to console, append message to log file
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
log error message

.Description
Send log to host as Error message and add content to logfile

.Parameter Exception
Exception object, used to get error message

.Parameter InvocationInfo
Invocation object info, use to get script name and line

.Parameter Title
Title to insert before content message

.Parameter LineAfterTitle
Insert blank line between title and message content

.Example
Write-LogError $($_.Exception) $($_.InvocationInfo)

.Outputs
Print message to console, append message to log file
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
			[string]$text = "at $(Get-Date -Format 'HH.mm.ss.fff') an $($Exception.GetType()) exception was thrown :$([Environment]::NewLine)"

			$message = ($InvocationInfo | Select-Object MyCommand, ScriptLineNumber, ScriptName, Line, OffsetInLine | Format-List | Out-String)
			$message += [Environment]::NewLine

			if ($Exception -is [System.Reflection.ReflectionTypeLoadException])
			{
				$message += ($Exception | Select-Object Message, Source, LoaderExceptions, StackTrace | Format-List | Out-String)
			}
			else
			{
				$message += ($Exception | Select-Object Message, Source, ErrorCode, InnerException, StackTrace | Format-List | Out-String)
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
Trace event message

.Description
Send log message to windows events log

.Parameter Message
Content of the message to log

.Parameter Type
Message type as string : Information, Error, Warning, SuccessAudit, FailureAudit

.Parameter Source
Event source (will be created if not exist)

.Parameter LogName
Destination log : Application, Security, Setup, System, custom log name..

.Parameter EventID
ID of the event

.Parameter Computer
Destination computer

.Example
Write-LogEvent "bonjour"

.Example
Write-LogEvent "bonjour" "Warning" "Batch"

.Outputs
Append message to event log
#>
Function Write-LogEvent
{
	[CmdletBinding()]
	Param ( [Parameter(Mandatory=$true,Position=0)][string]$Message,
			[Parameter(Mandatory=$false,Position=1)][string]$Type="Information",
			[Parameter(Mandatory=$false,Position=2)][string]$Source="tfsKit",
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

		Write-EventLog -LogName $LogName -Source $Source -ComputerName $ComputerName -EventId $EventID -Message $Message -EntryType $Type
	}
	catch
	{
		# log warning
		Write-LogWarning "Error Write-LogEvent -Exception $($_.Exception.Message)"
	}
}

<#
.Synopsis
Trace Host message

.Description
Send trace to host as Host message and add content to logfile

.Parameter Message
Content of the message to trace

.Parameter Title
Title to insert before content message

.Parameter LineBefore
Insert blank line before text

.Parameter LineAfter
Insert blank line after text

.Parameter LineAfterTitle
Insert blank line between title and message content

.Example
Write-LogHost "bonjour"

.Outputs
Print message to console, append message to log file
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
Log Object after formating display

.Description
Format Object into a proper readable format

.Parameter Object
Object to log

.Parameter Title
Add a title to the object report

.Parameter Level
Set log level for message output ("Host","Verbose","Warning","Debug","Error")

.Parameter Format
Object output formatting ("List","Table","Wide")

.Parameter SkipLineBefore
Do not Insert blank line before object

.Parameter SkipLineAfter
Do not Insert blank line after object

.Parameter LineAfterTitle
Insert blank line between title and message content

.Example
Write-LogObject $object

.Example
Write-LogObject $object -Format "Table"

.Outputs
Print message to console, append message to log file
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
Trace Verbose message

.Description
Send trace to host as Verbose message and add content to logfile

.Parameter Message
Content of the message to trace

.Parameter Title
Title to insert before content message

.Parameter LineBefore
Insert blank line before text

.Parameter LineAfter
Insert blank line after text

.Parameter LineAfterTitle
Insert blank line between title and message content

.Example
Write-LogVerbose "bonjour"

.Outputs
Print message to console, append message to log file
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
Log Warning message

.Description
Send log to host as Warning message and add content to logfile

.Parameter Message
Content of the message to trace

.Parameter Title
Title to insert before content message

.Parameter LineBefore
Insert blank line before text

.Parameter LineAfter
Insert blank line after text

.Parameter LineAfterTitle
Insert blank line between title and message content

.Example
Write-LogWarning "bonjour"

.Outputs
Print message to console, append message to log file
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

#endregion

#region XML

<#
.Synopsis
Return xml data from file

.Description
Iterate short names array then test-path, complete short name, set path, read and return xml file as object

.Parameter File
File name without extension, or file full path if you use switch -IsFullPath

.Parameter Folder
File folder if file is not hosted in Root\Data folder (not mandatory)

.Parameter IsFullPath
Switch set to true if $File is a file full path

.Example
Get-Data "CodeAnalysis"
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

		$returnObject = [xml](Get-Content $filePath -Encoding UTF8)

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

<#
.Synopsis
Get Xml element or attribute value

.Description
Get value from an xml element or attribute identified by XPath

.Parameter File
Full path to a xml file as string, or xml elements as XmlElements

.Parameter XPath
Xpath to property to get

.Example
Get-XPathValue "C:\example.xml" "/system.web/authentication/@mode"

.Outputs
Element value on success / null on fail
#>
Function Get-XPathValue
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][object]$File,
			[Parameter(Mandatory=$true,Position=1)][string]$XPath )

    Write-LogDebug "Start Get-XPathValue"

    #initialize return value
	$returnValue = $null

    try
    {
		Write-LogVerbose "opening $File to get $XPath"
		
		# get xml
		if ($File -is [string])
		{
			$xmlElements = [xml](Get-Content $File)
		}
		elseif ($File -is [Xml.XmlNode])
		{
			$xmlElements = $File
		}

		# get node corresponding to XPath
		$xmlNode = $xmlElements.SelectSingleNode($XPath)
        
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
Set / Add / Remove / Replace / Comment or UnComment an xml element or attribute identified by XPath

.Parameter File
Full path to a xml file as string, or xml elements as XmlElements

.Parameter XPath
Xpath to the edited element or attribute

.Parameter Value
Value to Set \ Add \ Replace

.Parameter Type
Edition type : setValue, add, remove, replace, comment, unComment

.Parameter addMode
Insertion mode , insert Value before or after Xpath

.Example
Set-XPathValue "C:\example.xml" "/system.web/authentication/@mode" "true"

.Example
Set-XPathValue "C:\example.xml" "/system.web/authentication/@mode" -Type "comment"

.Outputs
True on success / False on fail
#>
Function Set-XPathValue
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][object]$File,
			[Parameter(Mandatory=$true,Position=1)][string]$XPath,
			[Parameter(Mandatory=$false,Position=2)][AllowNull()][string]$Value,
			[Parameter(Mandatory=$false,Position=3)][AllowNull()][ValidateSet("","setValue", "add", "remove", "replace", "comment", "unComment")][string]$Type="setValue",
			[Parameter(Mandatory=$false,Position=4)][AllowNull()][ValidateSet("","before", "after", "in")][string]$AddMode="in")

    Write-LogDebug "Start Set-XPathValue"

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

		# get xml
		if ($File -is [string])
		{
			Write-LogVerbose "editing file $File $message"
			
			$xmlElements = [xml](Get-Content $File)
		}
		elseif ($File -is [Xml.XmlNode])
		{
			Write-LogVerbose "editing xml elements $message"
					
			$xmlElements = $File
		}

		# get node corresponding to XPath
		if ($Type -ne "unComment")
		{
			$xmlNode = $xmlElements.SelectSingleNode($XPath)
		}
		else
		{
			$parentXPath = Get-InString $XPath '/' -Last
			$childXPath = $XPath -replace "$parentXPath/", ''

			$xmlNode = $xmlElements.SelectNodes("$parentXPath/comment()") | ? { $_.InnerText -match "<$childXPath" }
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

				$tempNode = $xmlElements.CreateElement("tempNode")
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

				$tempNode = $xmlElements.CreateElement("tempNode")
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

				$xmlComment = $xmlElements.CreateComment($xmlNode.OuterXml)
				$xmlNode.ParentNode.ReplaceChild($xmlComment, $xmlNode)
			}
			"unComment"
			{
				Write-LogVerbose "unComment $XPath"

				$tempNode = $xmlElements.CreateElement("tempNode")
				$tempNode.InnerXml = $xmlNode.Value

				$xmlNode.ParentNode.ReplaceChild($($tempNode.SelectSingleNode("/$childXPath")), $xmlNode)
			}
		}

		# save file or assign modified Xml.Elements to returnValue
		if ($File -is [string])
		{
			$xmlElements.Save($File)
			
			[boolean]$returnValue = $true
		}
		elseif ($File -is [Xml.XmlNode])
		{
			[Xml.XmlNode]$returnValue = $xmlElements
		}

	    # trace Success
	    Write-LogDebug "Success Set-XPathValue"
    }
    catch
    {
        [boolean]$returnValue = $false

        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }

    return $returnValue
}

<#
.Synopsis
Test if an attribute or element exist

.Description
Send true if XPath is found on File, false if not

.Parameter File
Full path to a xml file as string, or xml elements as XmlElements

.Parameter XPath
Xpath to test

.Example
Test-XPath "C:\example.xml" "/system.web/authentication/@mode"

.Outputs
True if XPath exist / False if XPath does not exist or method fail
#>
Function Test-XPath
{
    [CmdletBinding()]
    Param (	[Parameter(Mandatory=$true,Position=0)][object]$File,
			[Parameter(Mandatory=$true,Position=1)][string]$XPath )

    Write-LogDebug "Start Test-XPath"

    #initialize return value
	$returnValue = $false

    try
    {
		Write-LogVerbose "opening $File to test if $XPath exists"

		# get xml
		if ($File -is [string])
		{
			$xmlElements = [xml](Get-Content $File)
		}
		elseif ($File -is [Xml.XmlNode])
		{
			$xmlElements = $File
		}

		# get node corresponding to XPath
		$xmlNode = $xmlElements.SelectSingleNode($XPath)
        
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

#region IO

<#
.Synopsis
Import assemblies array from ColorKit
 
.Description
Import assemblies by assembly short name array
 
.Parameter ShortNames
Short names list of the assemblies to import

.Parameter NotMicrosoft
Switch : select if you do not want to add "Microsoft" before assembly short name

.Example
Add-Assemblies @("SqlServer.SqlEnum", "SqlServer.Smo")
#>
Function Add-Assemblies
{
    [CmdletBinding()]
    Param
    (   [Parameter(Mandatory=$false,Position=0)][array]$ShortNames,
		[Parameter(Mandatory=$false,Position=1)][switch]$NotMicrosoft )

    Write-LogDebug "Start Add-Assemblies"

    try
    {
        foreach ($name in $ShortNames)
        {
			if ($NotMicrosoft)
			{
				$assemblyName = "$name.dll"
			}
			else
			{
				$assemblyName = "Microsoft.$name.dll"
			}

            $assemblyPath = [IO.Path]::Combine($global:AssembliesFolder, $assemblyName)
			
			Write-LogVerbose "Add-Type -Path $assemblyPath"
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
Import modules from Kit Modules folder
 
.Description
Import modules by module short name with -Global & -Force switch
 
.Parameter Names
Short names list of the modules to import (not mandatory, default = all modules)

.Parameter Assert
Switch : select if you want to test already loaded modules

.Example
Add-Modules "Infra"
#>
Function Add-Modules
{
    [CmdletBinding()]
    Param
    (   [Parameter(Mandatory=$false,Position=0)][array]$ShortNames,
		[Parameter(Mandatory=$false,Position=1)][switch]$Assert )
	
	Write-LogDebug "Start Add-Modules"

    try
    {
        if ($ShortNames)
        {
            foreach ($shortName in $ShortNames)
            {
				if (!$Assert -or !(Get-Module -Name "Module-$shortName"))
				{
					$modulePath = [IO.Path]::Combine($global:ModulesFolder, "Module-$shortName.psm1")
					Import-Module $modulePath -Force -Global
				}
            }
        }
        else
        {
			if ($Assert)
			{
				Get-ChildItem ($global:ModulesFolder) | ? {!(Get-Module -Name $_.BaseName)} | % { Write-LogVerbose "Import-Module -Name $_.FullName -Force -Global"; Import-Module -Name $_.FullName -Force -Global }
			}
			else
			{
				Get-ChildItem ($global:ModulesFolder) | % { Write-LogVerbose "Import-Module -Name $_.FullName -Force -Global"; Import-Module -Name $_.FullName -Force -Global }
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
Create folders if they don't not exists

.Description
Create folders if they don't exists, using [IO.Directory]::Exists and New-Item -Force

.Parameter Folders
Full path list of the folders to assert

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
String extract fonction

.Description
Get string part delimited by a separator

.Parameter InString
String to seek

.Parameter Delimiter
Characters delimiter

.Parameter SecondDelimiter
Characters second delimiter

.Parameter Last
Switch, if true get last occurence of delimiter (default = first)

.Parameter SecondLast
Switch, if true get last occurence of second delimiter (default = first)

.Parameter Right
Switch, if true get right part of the string (default = left, switch disabled if 2 delimiters)

.Parameter Up
Switch, if true get uppercase result

.Parameter Low
Switch, if true get lowercase result

.Outputs
Searched string or blank if not found
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
	$returnValue = [string]::Empty

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
Prompt for user confirmation

.Description
Send prompt for a Yes\No question

.Parameter Title
Text for prompt title

.Parameter Message
Text for prompt message

.Parameter YesCaption
Text for yes option

.Parameter NoCaption
Text for no caption

.Example
Get-Prompt "proceed installation ?"

.Outputs
True or False depending on user choice
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
Convert secure string to string

.Description
Convert secure string to plain text string using Runtime.InteropServices.Marshal library

.Parameter SecureString
Secure string to convert

.Example
Get-SecureString -SecureString $secureStr

.Outputs
Secure string as plain text string
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
Get sha256 hash from file path

.Description
Perform a sha256 compute then convert to String and return

.Parameter File
File path to the computed file

.Outputs
File SHA256 hash as base64 String \ null if exception
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
TempFolder path on success, null on method fail
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
Remove folder

.Description
Deals with remove-item bug on recurse + not empty folders

.Parameter FolderPath
Full path to the folder to remove

.Example
Remove-Folder "C:\Test"
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
Convert string to secure string

.Description
Convert string to secure string using ConvertTo-String with -AsPlainText & -Force switchs

.Parameter String
String to convert

.Example
Set-SecureString "bonjour"

.Outputs
String as secure string
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

#region WEB

<#
.Synopsis
get file from http

.Description
download file from http, then perform sha256 checksum integrity

.Parameter FileSource
http source file to download

.Parameter FileDestination
destination file (optionnal, null or default = ScriptPath\HttpFileName)

.Parameter SHA256
sha256 hash of the file, converted to Base64 String (optionnal, null or default = no verification)

.Example
Get-HttpFile -FileSource "http://blob/test.zip" -FileDestination "C:\tools\WelcomeApp.zip" -SHA256 "YDFGKJH5KJHLKJH123UUJH=="

.Output
true (download and checksum success) or false (download or checksum fail)
#>
Function Get-HttpFile
{
    Param ( [parameter(Mandatory=$true,Position=0)][string]$FileSource,
            [parameter(Mandatory=$false,Position=1)][string]$FileDestination="$PSScriptRoot\$(Get-InString $FileSource "/" -Last -Right)",
            [parameter(Mandatory=$false,Position=2)][string]$SHA256 )

	Write-LogDebug "Start Get-HttpFile"
				
    try
    {
		Write-LogVerbose "downloading file $FileSource"
			
		# initialize returnValue
		$returnValue = $false
							
		# Create destination tree
		New-Item -Path $FileDestination -Type "file" -Force
		
        # set webclient, download file, dispose webClient
        $webclient = New-Object System.Net.WebClient						
        $webclient.DownloadFile($FileSource, $FileDestination)
        $webclient.Dispose()

        # test if destination file is present
        if (Test-Path $FileDestination)
        {
            # perform sha256 checksum to check file authenticity
            if ($SHA256)
            {
                # get sha256 for local file
                $sha256Local = $(Get-SHA256 $FileDestination)
				
				Write-LogVerbose "comparing file checkum, local: $sha256Local ; server: $SHA256"
				
                # test sha256 are equals
                if ($SHA256 -eq $sha256Local)
                {   
					Write-LogHost "SHA256 local and distant are identical"				
                    $returnValue = $true
                }
				else
				{
					Write-LogWarning "SHA256 local and distant are different"	
				}
            }
            else
            {
				Write-LogVerbose "no checksum file comparaison set"
				
				$returnValue = $true 
            }
        }
        else
        {
            Throw "destination file $FileDestination is missing"
        }
		
		# trace Success
		Write-LogDebug "Success Get-SHA256"
    }
    catch
    {
        # log error
        Write-LogError $($_.Exception) $($_.InvocationInfo) 
    }
	
	return $returnValue
}

<#
.Synopsis
Get data from web services

.Description
Get data from web services using Msxml2.XMLHTTP

.Parameter Url
web svc url to query

.Parameter Headers
Request headers array (optionnal, default : no headers)

.Parameter Path
Xpath selection from response (optionnal, default : full document)

.Output
Xml result on success, null on fail
#>
Function Get-WebService
{
	[CmdletBinding()]
	Param (	[Parameter(Mandatory=$true,Position=0)][string]$Url,
            [Parameter(Mandatory=$false,Position=1)][array]$Headers,
            [Parameter(Mandatory=$false,Position=2)][string]$XPath )

	Write-LogDebug "Start Get-WebService"
				
    try
    {
		Write-LogVerbose "Getting web service $Url"
		
		Write-LogObject $Headers "Headers"
					
		# create http request object
		$HttpReq = New-Object -ComObject Msxml2.XMLHTTP
		
		# set http request URL
		$HttpReq.open('GET', $Url, $false)
		
        if ($Headers)
        {
            # set http request headers
		    foreach($Header in $Headers)
		    {
			    $HttpReq.SetRequestHeader($Header[0], $Header[1])
		    }
        }

		# send http request
		$HttpReq.Send()
        
		# get http request response
		$xmlResponse = $HttpReq.ResponseXML
		
		# Assign Node text to ReturnValue
        if ($XPath)
        {
            $xmlResult = Get-XPathValue $($xmlResponse.Xml) $XPath
        }
        else
        {
            $xmlResult += $xmlResponse.Xml
        }	

		# trace Success
		Write-LogDebug "Success Get-WebService"
    }
    catch
    {
		$xmlResult = $null
		
        # log Error
        Write-LogError $($_.Exception) $($_.InvocationInfo)
    }
	
	return $xmlResult
}

<#
.Synopsis
Write a Object array or Hashtable/OrderedDictionary collection to html table

.Description
Write a one level object array or Hashtable/OrderedDictionary collection into html table

.Parameter Object
One level object array or Hashtable/OrderedDictionary collection to report

.Parameter TableClass
If set, define the table class

.Parameter ThClass
If set, define the table header class

.Parameter TrClass
If set, define the table row class

.Parameter TdClass
If set, define the table data class

.Parameter Horizontal
Switch, set to true if you want to print html table horizontaly

.Parameter SkipTh
Switch, set to true if you don't want to print table header

.Example
Write-ObjectToHtml $object -FirstLineAsTh -ThClass "bigTitle"

.Outputs
object formated as HTML table on success / null on fail
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
		    $htmlTable = [string]::Empty
        
            if ($TableClass) { $TableClass = " class=`"$TableClass`"" } else { $TableClass = [string]::Empty }
            if ($ThClass) { $ThClass = " class=`"$ThClass`"" } else { $ThClass = [string]::Empty }
            if ($TrClass) { $TrClass = " class=`"$TrClass`"" } else { $TrClass = [string]::Empty }
            if ($TdClass) { $TdClass = " class=`"$TdClass`"" } else { $TdClass = [string]::Empty }
            if ($TableId) { $TableId = " Id=`"$TableId`"" } else { $TableId = [string]::Empty }

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
            $htmlTable = [string]::Empty
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

#end