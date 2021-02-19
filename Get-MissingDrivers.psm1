# Documentation home: https://github.com/engrit-illinois/Get-MissingDrivers
# By mseng3

function Get-MissingDrivers {
	
	param(
				
		[Parameter(Position=0,Mandatory=$true)]
		[string[]]$Computers,
		
		[string]$OUDN = "OU=Desktops,OU=Engineering,OU=Urbana,DC=ad,DC=uillinois,DC=edu",
		
		# ":ENGRIT:" will be replaced with "c:\engrit\logs\$($MODULE_NAME)_:TS:.csv"
		# ":TS:" will be replaced with start timestamp
		[string]$Csv,
		
		# ":ENGRIT:" will be replaced with "c:\engrit\logs\$($MODULE_NAME)_:TS:.log"
		# ":TS:" will be replaced with start timestamp
		[string]$Log,
		
		[switch]$NoConsoleOutput,
		[string]$Indent = "    ",
		[string]$LogFileTimestampFormat = "yyyy-MM-dd_HH-mm-ss",
		[string]$LogLineTimestampFormat = "[HH:mm:ss] ",
		#[string]$LogLineTimestampFormat = "[yyyy-MM-dd HH:mm:ss:ffff] ",
		#[string]$LogLineTimestampFormat = $null, # For no timestamp
		[int]$Verbosity = 0,
		
		
		[switch]$ReturnObject,
		
		[int]$CIMTimeoutSec = 60,
		
		# Maximum number of asynchronous jobs allowed to run at once
		# This is to limit how many computers are simultaneously being polled for profile data over the network
		# Set to 0 to disable the asynchronous mechanism entirely and just do it sequentially
		[int]$MaxAsyncJobs = 50
	)
	
	$MODULE_NAME = "Get-MissingDrivers"
	$ASYNC_JOBNAME_BASE = "Job_$($MODULE_NAME)"
	
	$ENGRIT_LOG_DIR = "c:\engrit\logs"
	$ENGRIT_LOG_FILENAME = "$($MODULE_NAME)_:TS:"
	
	$START_TIMESTAMP = Get-Date -Format $LogFileTimestampFormat
	if($Log) {
		$Log = $Log.Replace(":ENGRIT:","$($ENGRIT_LOG_DIR)\$($ENGRIT_LOG_FILENAME).log")
		$Log = $Log.Replace(":TS:",$START_TIMESTAMP)
	}
	if($Csv) {
		$Csv = $Csv.Replace(":ENGRIT:","$($ENGRIT_LOG_DIR)\$($ENGRIT_LOG_FILENAME).csv")
		$Csv = $Csv.Replace(":TS:",$START_TIMESTAMP)
	}
	
	$THING = "driver"
	$THINGS = "drivers"
	$THING_PROPERTY = "_Drivers"
	
	function log {
		param (
			[Parameter(Position=0)]
			[string]$Msg = "",

			[int]$L = 0, # level of indentation
			[int]$V = 0, # verbosity level
			
			[ValidateScript({[System.Enum]::GetValues([System.ConsoleColor]) -contains $_})]
			[string]$FC = (get-host).ui.rawui.ForegroundColor, # foreground color
			[ValidateScript({[System.Enum]::GetValues([System.ConsoleColor]) -contains $_})]
			[string]$BC = (get-host).ui.rawui.BackgroundColor, # background color
			
			[switch]$E, # error
			[switch]$NoTS, # omit timestamp
			[switch]$NoNL, # omit newline after output
			[switch]$NoConsole, # skip outputting to console
			[switch]$NoLog # skip logging to file
		)
			
		if($E) { $FC = "Red" }
		
		# Custom indent per message, good for making output much more readable
		for($i = 0; $i -lt $L; $i += 1) {
			$Msg = "$Indent$Msg"
		}
		
		# Add timestamp to each message
		# $NoTS parameter useful for making things like tables look cleaner
		if(!$NoTS) {
			if($LogLineTimestampFormat) {
				$ts = Get-Date -Format $LogLineTimestampFormat
			}
			$Msg = "$ts$Msg"
		}

		# Each message can be given a custom verbosity ($V), and so can be displayed or ignored depending on $Verbosity
		# Check if this particular message is too verbose for the given $Verbosity level
		if($V -le $Verbosity) {
		
			# Check if this particular message is supposed to be output to console
			if(!$NoConsole) {

				# Uncomment one of these depending on whether output goes to the console by default or not, such that the user can override the default
				#if($ConsoleOutput) {
				if(!$NoConsoleOutput)
				
					# If we're allowing console output, then Write-Host
					if($NoNL) {
						Write-Host $Msg -NoNewline -ForegroundColor $FC -BackgroundColor $BC
					}
					else {
						Write-Host $Msg -ForegroundColor $FC -BackgroundColor $BC
					}
				}
			}

			# Check if this particular message is supposed to be logged
			if(!$NoLog) {
				
				# If $Log was specified, then log to file
				if($Log) {
					
					# Check that the logfile already exists, and if not, then create it (and the full directory path that should contain it)
					if(!(Test-Path -PathType "Leaf" -Path $Log)) {
						New-Item -ItemType "File" -Force -Path $Log | Out-Null
					}

					if($NoNL) {
						$Msg | Out-File $Log -Append -NoNewline
					}
					else {
						$Msg | Out-File $Log -Append
					}
				}
			}
		}
	}
	
	function Log-Error($e, $L) {
		log "$($e.Exception.Message)" -L $l
		log "$($e.InvocationInfo.PositionMessage.Split("`n")[0])" -L ($L + 1)
	}
	
	function Get-CompNameString($comps) {
		$list = ""
		foreach($comp in $comps) {
			$list = "$list, $($comp.Name)"
		}
		$list = $list.Substring(2,$list.length - 2) # Remove leading ", "
		$list
	}

	function Get-Comps($compNames) {
		log "Getting computer names..."
		
		$comps = @()
		foreach($name in @($compNames)) {
			$comp = Get-ADComputer -Filter "name -like '$name'" -SearchBase $OUDN
			$comps += $comp
		}
		$list = Get-CompNameString $comps
		log "Found $($comps.count) computers in given array: $list." -L 1
	
		log "Done getting computer names." -V 2
		$comps
	}
	
	function AsyncGet-DataFrom($comp) {
		param(
			$comp,
			$CIMTimeoutSec,
			$THINGS,
			$THING_PROPERTY
		)
		
		# Note: cannot do any logging to the console from jobs (i.e. log()), because jobs run in another process.
		# Might be able to capture this with a bunch of extra code checking each job while it's still running, but that's totally not worth it for short jobs
		# Jobs might still be able to write to the log file, but it would be out of order and might cause race conditions with different processes accessing the same log file at the same time.
		# Also, any external functions and variable must be passed into the job since it's a separate process. For functions which call other functions and variables, this quickly becomes infeasible.
		# For simple scripts like this, I'll have to accept not getting any real-time feedback from async jobs.
		
		function log($msg) {
			# Any output from a job, even Write-Host, will only show up when you Receive-Job
			#Write-Host $msg
		}
	
		function Log-Error($e) {
			log "$($e.Exception.Message)"
			log "$($e.InvocationInfo.PositionMessage.Split("`n")[0])"
		}
			
		$compName = $comp.Name
		log "Getting $THINGS from `"$compName`"..."
		
		# Note to self: $error is a reserved variable name
		# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables
		try {
			# -ErrorAction Stop required for catch{} to get anything, otherwise Get-WmiObject and Get-CimInstance do not throw terminating errors by default
			# https://stackoverflow.com/questions/1142211/try-catch-does-not-seem-to-have-an-effect
			$data = Get-CIMInstance -ComputerName $compName -ClassName "Win32_PNPEntity" -OperationTimeoutSec $CIMTimeoutSec -ErrorAction "Stop"
		}
		catch {
			log "Error calling Get-CIMInstance on computer `"$compname`"!"
			Log-Error $_
			$errorMsg = $_.Exception.Message
		}
		if($errorMsg) {
			$comp | Add-Member -NotePropertyName "_Error" -NotePropertyValue $errorMsg -Force
		}
		
		log "Found $(@($data).count) $($THINGS)."
		if($data) {
			$comp | Add-Member -NotePropertyName $THING_PROPERTY -NotePropertyValue $data -Force
		}
		
		log "Done getting $THINGS from `"$compname`"."
		$comp
	}
	
	function Start-AsyncGetDataFrom($comp) {
		# If there are already the max number of jobs running, then wait
		$running = @(Get-Job | Where { $_.State -eq 'Running' })
		if($running.Count -ge $MaxAsyncJobs) {
			$running | Wait-Job -Any | Out-Null
		}
		
		# After waiting, start the job
		# Each job gets $THINGS, and returns a modified $comp object with the $THINGS included
		# We'll collect each new $comp object into the $comps array when we use Recieve-Job
		
		<# Turns out that using script-level functions in Start-Job ScriptBlocks is not that simple:
		
		#$comp = GetDriversFrom $comp
		#return $comp
		
		# https://stackoverflow.com/questions/7162090/how-do-i-start-a-job-of-a-function-i-just-defined
		# https://social.technet.microsoft.com/Forums/windowsserver/en-US/b68c1c68-e0f0-47b7-ba9f-749d06621a2c/calling-a-function-using-startjob?forum=winserverpowershell
		# https://stuart-moore.com/calling-a-powershell-function-in-a-start-job-script-block-when-its-defined-in-the-same-script/
		# https://stackoverflow.com/questions/15520404/how-to-call-a-powershell-function-within-the-script-from-start-job
		# https://stackoverflow.com/questions/8489060/reference-command-name-with-dashes
		# https://powershell.org/forums/topic/what-is-function-variable/
		# https://stackoverflow.com/questions/8489060/reference-command-name-with-dashes
		# https://stackoverflow.com/a/8491780/994622
		#>
		
		$scriptBlock = Get-Content function:AsyncGet-DataFrom
		
		$ts = Get-Date -Format "yyyy-MM-dd_HH-mm-ss-ffff"
		$compName = $comp.name
		$uniqueJobName = "$($ASYNC_JOBNAME_BASE)_$($compName)_$($ts)"
		Start-Job -Name $uniqueJobName -ArgumentList $comp,$CIMTimeoutSec,$THINGS,$THING_PROPERTY -ScriptBlock $scriptBlock | Out-Null
	}
	
	function AsyncGet-Data($comps) {
		# Async example: https://stackoverflow.com/a/24272099/994622
		
		# For each computer start an asynchronous job
		log "Starting async jobs to get $THINGS from computers..." -L 1
		log "-MaxAsyncJobs is set to $($MaxAsyncJobs)." -L 2
		$count = 0
		foreach ($comp in $comps) {
			log $comp.Name -L 2
			Start-AsyncGetDataFrom $comp
			$count += 1
		}
		log "Started $count jobs." -L 1
		
		# Wait for all the jobs to finish
		$jobNameQuery = "$($ASYNC_JOBNAME_BASE)_*"
		log "Waiting for async jobs to finish..." -L 1
		Get-Job -Name $jobNameQuery | Wait-Job | Out-Null

		# Once all jobs are done, start processing their output
		# We can't directly write over each $comp in $comps, because we don't know which one is which without doing a bunch of extra logic
		# So just make a new $comps instead
		$newComps = @()
		
		log "Receiving jobs..." -L 1
		$count = 0
		foreach($job in Receive-Job -Name $jobNameQuery) {
			$comp = $job
			log "Received job for computer `"$($comp.Name)`"." -L 2
			$newComps += $comp
			$count += 1
		}
		log "Received $count jobs." -L 1
		
		log "Removing jobs..." -L 1
		Remove-Job -Name $jobNameQuery
		
		$newComps
	}
	
	function Get-Data($comps) {
		log "Retrieving $($THINGS)..."
		
		if($MaxAsyncJobs -lt 1) {
			foreach($comp in $comps) {
				$comp = Get-DataFrom $comp
			}
		}
		else {
			$comps = AsyncGet-Data $comps
		}
		
		log "Done retrieving $($THINGS)." -V 2
		$comps
	}
	
	function Munge-Data($comps) {
		log "Munging data..."
		
		foreach($comp in $comps) {
			$compName = $comp.Name
			log "$compName" -L 1 -V 1
			
			foreach($item in $comp.$THING_PROPERTY) {
				# Do some work on each item
				#$customItemData = ($item.Name).Replace("this","that")
				
				# Save some custom data about this item
				#$item | Add-Member -NotePropertyName "_CustomItemData" -NotePropertyValue $customItemData -Force
			}
			
			# This gets me EVERY FLIPPIN TIME:
			# https://stackoverflow.com/questions/32919541/why-does-add-member-think-every-possible-property-already-exists-on-a-microsoft
			
			# Save some custom data discovered about this comp
			#$customCompData = @(($comp.$THING_PROPERTY)._CustomItemData).count
			#$comp | Add-Member -NotePropertyName "_CustomCompData" -NotePropertyValue $customCompData -Force
			
			log "Done with `"$compName`"." -L 1 -V 2
		}
		
		log "Done munging data." -V 2
		$comps
	}
	
	function Print-Data($comps) {
		log "Summary of $THINGS from all computers:"
		log ($comps | Format-Table | Out-String) -NoTS
	}
	
	function Export-Data($comps) {
		if($Csv) {
			log "-Csv was specified. Exporting data to `"$Csv`"..."
			
			$comps | Export-Csv -NoTypeInformation -Encoding "Ascii" -Path $Csv
		}
	}
	
	function Return-Object($comps) {
		if($ReturnObject) {
			$comps
		}
	}
	
	function Get-FlatData($comps) {
		$flatData = @()
		foreach($comp in $comps) {
			$flatData += @($comp.$THING_PROPERTY)
		}
		$flatData
	}
	
	function Get-RunTime($startTime) {
		$endTime = Get-Date
		$runTime = New-TimeSpan -Start $startTime -End $endTime
		$runTime
	}

	function Do-Stuff {
		$startTime = Get-Date
		
		$comps = Get-Comps $Computers
		$comps = Get-Data $comps
		$comps = Munge-Data $comps
		
		$flatData = Get-FlatData $comps
		
		Print-Data $comps
		Export-Data $comps
		Return-Object $comps
		
		$runTime = Get-RunTime $startTime
		log "Runtime: $runTime"
	}
	
	Do-Stuff

	log "EOF"

}