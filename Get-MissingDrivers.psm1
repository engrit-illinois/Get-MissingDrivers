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
		
		[switch]$IncludeValidDrivers,
		
		[ValidateSet("FlatDataSummary","FlatData","Computers")]
		[string]$OutputFormat = "FlatDataSummary",
		
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
				if(!$NoConsoleOutput) {
				
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
						log "Logging to `"$Log`"."
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
	
	function AsyncGet-DataFrom {
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
			
			# https://social.technet.microsoft.com/Forums/en-US/54c4c520-2831-4f7f-9fab-a32653a61cac/find-unknown-devices-with-powershell?forum=winserverpowershell
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
			
			$cadStatus = 0
			foreach($item in $comp.$THING_PROPERTY) {
				$code = $item.ConfigManagerErrorCode
				if($code -ne 0) { $badStatusCount += 1 }
				$status = Translate-ConfigManagerErrorCode $code
				$item | Add-Member -NotePropertyName "_Status" -NotePropertyValue $status -Force
			}
			
			# This gets me EVERY FLIPPIN TIME:
			# https://stackoverflow.com/questions/32919541/why-does-add-member-think-every-possible-property-already-exists-on-a-microsoft
			
			# Save some custom data discovered about this comp
			$comp | Add-Member -NotePropertyName "_BadStatusCount" -NotePropertyValue $errorCount -Force
			
			log "Done with `"$compName`"." -L 1 -V 2
		}
		
		log "Done munging data." -V 2
		$comps
	}
	
	# https://social.technet.microsoft.com/Forums/en-US/54c4c520-2831-4f7f-9fab-a32653a61cac/find-unknown-devices-with-powershell?forum=winserverpowershell
	# https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-pnpentity
	# https://support.microsoft.com/en-us/topic/error-codes-in-device-manager-in-windows-524e9e89-4dee-8883-0afa-6bca0456324e
	# https://docs.microsoft.com/en-us/windows-hardware/drivers/install/device-manager-error-messages
	# Per the above docs, the error messages and their codes seem to have changed slightly and been added to at some point (between 2017 and 2021).
	# I can't be bothered to figure out which list is correct for which operating systems.
	# This script is almost exclusively written for finding drivers with error code 28, which hasn't changed.
	# I did add code 43, because I saw one instance of that during my testing.
	# I also saw 22 and 24 a couple times, which also haven't changed.
	function Translate-ConfigManagerErrorCode($code) {
		switch($code) {
			0 {"Device is working properly."}
			1 {"Device is not configured correctly."}
			2 {"Windows cannot load the driver for this device."}
			3 {"Driver for this device might be corrupted, or the system may be low on memory or other resources."}
			4 {"Device is not working properly. One of its drivers or the registry might be corrupted."}
			5 {"Driver for the device requires a resource that Windows cannot manage."}
			6 {"Boot configuration for the device conflicts with other devices."}
			7 {"Cannot filter."}
			8 {"Driver loader for the device is missing."}
			9 {"Device is not working properly. The controlling firmware is incorrectly reporting the resources for the device."}
			10 {"Device cannot start."}
			11 {"Device failed."}
			12 {"Device cannot find enough free resources to use."}
			13 {"Windows cannot verify the device's resources."}
			14 {"Device cannot work properly until the computer is restarted."}
			15 {"Device is not working properly due to a possible re-enumeration problem."}
			16 {"Windows cannot identify all of the resources that the device uses."}
			17 {"Device is requesting an unknown resource type."}
			18 {"Device drivers must be reinstalled."}
			19 {"Failure using the VxD loader."}
			20 {"Registry might be corrupted."}
			21 {"System failure. If changing the device driver is ineffective, see the hardware documentation. Windows is removing the device."}
			22 {"Device is disabled."}
			23 {"System failure. If changing the device driver is ineffective, see the hardware documentation."}
			24 {"Device is not present, not working properly, or does not have all of its drivers installed."}
			25 {"Windows is still setting up the device."}
			26 {"Windows is still setting up the device."}
			27 {"Device does not have valid log configuration."}
			28 {"Device drivers are not installed."}
			29 {"Device is disabled. The device firmware did not provide the required resources."}
			30 {"Device is using an IRQ resource that another device is using."}
			31 {"Device is not working properly.  Windows cannot load the required device drivers."}
			43 {"Windows has stopped this device because it has reported problems."}
		}
	}
	
	function Prune-Data($comps) {
		$newComps = @()
		foreach($comp in $comps) {
			# Get rid of computers which didn't return any data
			if($comp.$THING_PROPERTY) {
				# Get rid of valid drivers
				$newComp = $comp
				
				if(!$IncludeValidDrivers) {
					$newComp.$THING_PROPERTY = $comp.$THING_PROPERTY | Where { $_.ConfigManagerErrorCode -ne 0 }
				}
				$newComps += @($newComp)
			}
		}
		$newComps
	}
	function Get-SummaryData($data) {
		$data | Select PSComputerName,@{Name="_Error",Expression={$_.ConfigManagerErrorCode}},_Status,Name,DeviceID
	}
	
	function Print-Data($data) {
		log "Summary of $THINGS from all computers:"
		log "" -NoTS
		log ($data | Format-Table | Out-String).Trim() -NoTS
		log "" -NoTS
	}
	
	function Export-Data($data) {
		if($Csv) {
			log "-Csv was specified. Exporting data to `"$Csv`"..."
			
			$data | Export-Csv -NoTypeInformation -Encoding "Ascii" -Path $Csv
		}
	}
	
	function Return-Object($data) {
		if($ReturnObject) {
			$data
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
		
		$comps = Prune-Data $comps
		
		if($OutputFormat -eq "Computers") {
			$data = $comps
		}
		else {
			$data = Get-FlatData $comps
			if($OutputFormat -eq "FlatDataSummary") {
				$data = Get-SummaryData $data
			}
		}
		
		Print-Data $data
		Export-Data $data
		Return-Object $data
		
		$runTime = Get-RunTime $startTime
		log "Runtime: $runTime"
	}
	
	Do-Stuff

	log "EOF"

}