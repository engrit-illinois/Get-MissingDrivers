# Summary

This script asynchronously polls an array of computers and reports whether they have any devices with missing or problematic drivers.  

# Usage

1. Download `Get-MissingDrivers.psm1`
2. Import it as a module: `Import-Module "c:\path\to\Get-MissingDrivers.psm1"`
3. Run it using the parameters documented below
- e.g. `Get-MissingDrivers -Computers "gelib-4c-*" -Csv ":ENGRIT:"

# Parameters

### -Computers \<string[]\>
Required string array.  
Array of computer names and/or computer name wildcard queries to poll.  
Computers must exist in AD, within the OU specified by `-OUDN`.  
e.g. `-Computers "gelib-4c-*","eh-406b1-01","mel-1001-*"`  

### -OUDN \<string\>
Optional string.  
The OU in which computers given by the value of `-Computers` must exist.  
Computers not found in this OU will be ignored.  
Default is `OU=Desktops,OU=Engineering,OU=Urbana,DC=ad,DC=uillinois,DC=edu`.  

### -Csv
Optional string.  
The full path of a file to export polled data to, in CSV format.  
If omitted, no CSV will be created.  
If `:TS:` is given as part of the string, it will be replaced by a timestamp of when the script was started, with a format specified by `-LogFileTimestampFormat`.  
Specify `:ENGRIT:` to use a default path (i.e. `c:\engrit\logs\Get-MissingDrivers_<timestamp>.csv`).  

### -Log
Optional string.  
The full path of a file to log to.  
If omitted, no log will be created.  
If `:TS:` is given as part of the string, it will be replaced by a timestamp of when the script was started, with a format specified by `-LogFileTimestampFormat`.  
Specify `:ENGRIT:` to use a default path (i.e. `c:\engrit\logs\Get-MissingDrivers_<timestamp>.log`).  

### -NoConsoleOutput
Optional switch.  
If specified, progress output is not logged to the console.  

### -Indent \<string\>
Optional string.  
The string used as an indent, when indenting log entries.  
Default is four space characters.  

### -LogFileTimestampFormat \<string\>
Optional string.  
The format of the timestamp used in filenames which include `:TS:`.  
Default is `yyyy-MM-dd_HH-mm-ss`.  

### -LogLineTimestampFormat \<string\>
Optional string.  
The format of the timestamp which prepends each log line.  
Default is `[HH:mm:ss]⎵`.  

### -Verbosity \<int\>
Optional integer.  
The level of verbosity to include in output logged to the console and logfile.  
Currently not significantly implemented.  
Default is `0`.  

### -IncludeValidDrivers
Optional switch.  
If specified, all driver data gathered from each computer is returned.  
By default, only driver items retrieved whose `ConfigManagerErrorCode` property contains a value not equal to `0` will be returned (hence "Get-_MISSING_Drivers").  

### -OutputFormat ["FlatDataSummary" | "FlatData" | "Computers"]
Optional string from a predefined set of strings.  
The format of the data output to the screen, CSV, and object.  
- Specifying `FlatDataSummary` returns an array of items gathered from the given computers, each with a `PSComputerName` property identifying which computer it came from. The data in each item is limited to only the most relevant information.
- Specifying `FlatData` returns an array of items, similar to `FlatDataSummary`, but all data for each item is returned.
- Specifying `Computers` returns an array of computers, which each have a `_Drivers` property containing the array of driver data gathered.
Default is `FlatDataSummary`.  

### -ReturnObject
Optional switch.  
If specified, the module returns an object to the pipeline, which contains all of the data gathered during execution.  

### -CIMTimeoutSec \<int\>
Optional integer.  
The number of seconds to wait before timing out `Get-CIMInstance` operations (the mechanism by which the script retrieves data from remote computers).  
Default is 60.  

### -MaxAsyncJobs \<int\>
Optional integer.  
The maximum number of asynchronous jobs allowed to be spawned simultaneously.  
The script spawns a unique asynchronous process for each computer that it will poll, which significantly cuts down the runtime.  
Default is `50`. This is to avoid the potential for network congestion and the possibility of the script being identified as malicious by antimalware processes and external network monitoring.  
To disable asynchronous jobs and external processes entirely, running everything sequentially in the same process, specify `0`. This will drastically increase runtime for large numbers of computers.  

# Sources
- Primarily based off of code from [this thread](https://social.technet.microsoft.com/Forums/en-US/54c4c520-2831-4f7f-9fab-a32653a61cac/find-unknown-devices-with-powershell?forum=winserverpowershell).
- Some docs which list the meanings of error codes:
  - https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-pnpentity
  - https://support.microsoft.com/en-us/topic/error-codes-in-device-manager-in-windows-524e9e89-4dee-8883-0afa-6bca0456324e
  - https://docs.microsoft.com/en-us/windows-hardware/drivers/install/device-manager-error-messages

# Notes
- Per the above docs, the error messages and their codes seem to have changed slightly and have been extended at some point (between 2017 and 2021).
  - Formerly there were only 32 codes (0 to 31). Now there appears to be up to 57, with some being changed or retired, depending on where you get your information.
  - This script currently uses the older definitions. I can't be bothered to figure out which codes are correct for which operating system versions. This script is written almost exclusively for the purpose of finding drivers with error code `28`, which hasn't changed in the updated definitions.
  - I did add code `43`, because I saw one instance of that during my testing. I also saw `22` and `24` a few times, which also haven't changed.
- By mseng3. See my other projects here: https://github.com/mmseng/code-compendium.
