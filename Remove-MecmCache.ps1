<#
    Name: 
        Remove-MecmCache

    Constituent source files: 
        1. Remove-MecmCache.ps1

    Descirption: Script for removing MecmCache on target computer name or .csv containing target computer names.
    
    Requirements: 
        1. Run Powershell as administrator
        2. Remote permissions for service control and delete at c:\windows\ccmcache

    Source Code Control (change control):
        <TBD>

    Created by:
    Andrew Goetz
    v.1
#>

<#
.SYNOPSIS
    Remove MECM cache on a target computer or list of computers in .csv file.
.DESCRIPTION
    1. Remove MECM cache on a target computer or list of computers in .csv file. 
    2. Ensure MECM service is stopped before files are deleted and restarted at end.
    3. Generate report at end.
.INPUTS
    -Target "computername"
    OR
    -Target "c:\path\to\something.csv"
.OUTPUTS
    This Cmdlet writes output to the Host.
.EXAMPLE
    1. Remove-MecmCache -Target "computername"
     - Start the cmdlet in single target mode. 
    2. Remove-MecmCache -Target "c:\path\to\something.csv"
     - Start the cmdlet in multiple target mode
    3. Remove-MecmCache "C:\Temp\targets.csv" | Tee-Object -FilePath "c:\temp\log.txt" -Append
     - Start the cmdlet in multiple target mode and output results to console as well as log.txt
.NOTES
    1. PowerShell must be executed as administrator.
    2. User must have permissions to start/stop services remotely and permission to delete from: c:\windows\ccmcache access on target computer(s).
#>
#ENDREGION

function Remove-MecmCache ([Parameter(Mandatory=$true)] $Target, $CachePath = "C:\Windows\ccmcache", $Service = "CcmExec")
{
    function Start-Routine ($Target,$CachePath,$Service )
    {
        try{
            Write-Log -Message "Starting Process on: $Target"
            $runner = Get-RunnerObject -TargetName $Target -Service $Service -CachePath $CachePath
            Get-FileType($runner)
            Get-Targets($runner)
            Start-PowerShellSessions($runner)
            Start-Actions($runner)
            Stop-PowerShellSessions($runner)
            Get-Report($runner)
            $runner.TimeEnded = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $runner.TimeElapsed = Get-ElapsedTime -StartTime $runner.TimeStarted -EndTime $runner.TimeEnded
            Write-Log -Message "Remove-MecmCache took: $($runner.TimeElapsed) to run."
            Write-Log -Message "Remove-MecmCache Completed."
        }catch{
            $msg = Get-ExceptionMessage($_)
            Write-Log("A critical exception has occurred in Start-Routine!")
            Write-Log($msg)
        }
    }

    function Start-Action($Object)
    {
        try{
            Get-FreeSpace -Object $Object
            Get-IsCacheFound -Object $Object
            if (!$Object.CacheFound)
            {
                $Object.Result = "Failed"
                $Object.ResultDetails = "CachePath: $($Object.CachePath) does not exist"
                throw($Object.ResultDetails)
            }
            Get-FilesAtPath($Object)
            if ($Object.CachedItemCount -gt 0){
                Stop-RemoteService $Object
                if ($Object.ServiceStatus -ne "Stopped")  
                {
                    $Object.Result = "Failed"
                    $Object.ResultDetails = "Service: $($Object.Service): failed to stop"
                    throw($Object.ResultDetails)
                }   
                Remove-Files $Object
                Start-RemoteService($Object)
                if ($Object.ServiceStatus -ne "Running"){
                    $Object.Result = "Failed"
                    $Object.ResultDetails = "Service: $($Object.Service) failed to start"
                    throw($Object.ResultDetails)
                }
                Get-FreeSpace -Object $Object
                $Object.Result = "Success"
            }
            else{
                $Object.ResultDetails = "FilePath: $($Object.CachePath) has no files"
                throw($Object.ResultDetails)
            }
        }catch{
            $msg = Get-ExceptionMessage($_)
            Write-Log("A critical exception has occurred!")
            Write-Log("$($msg)")
            $Object.Result = "Failed"
            $Object.ResultDetails = $msg
        }
    }

    function Start-Actions($Object)
    {
        foreach ($target in $Object.TargetItems)
        {
            if ($target.Result -ne "Failed")
            {
                try
                {
                    $target.Result = "Processing"
                    Start-Action($target)
                }
                catch
                {
                    $target.Result = "Failed"
                    $target.ResultDetails = $_.Message
                }
            }
        }
    }

    function Start-PowerShellSessions($Runner)
    {
        foreach ($item in $Runner.TargetItems)
        {
            Start-PowerShellSession($item)
        }
    }

    function Start-PowerShellSession($Target)
    {
        try{
            Write-Log("---Starting PowerShell session on: $($Target.ComputerName)...---")
            $Target.Session = New-PSSession -ComputerName $Target.ComputerName -ErrorAction Stop
            Write-Log("Started PowerShell session on: $($Target.ComputerName).---")
        }
        catch{
            Write-Log("CRITICAL: " + $_.Exception.Message)
            $message = Get-ExceptionMessage($_)
            $Target.ResultDetails = "CRITICAL: $message on $($Target.ComputerName)"
            $Target.Result = "Failed"
        }
    }

    function Stop-PowerShellSession($Target)
    {
        try{
            Remove-PSSession -Session $Target.Session -ErrorAction Stop
            Write-Log("---PowerShell session removed for: $($Target.ComputerName).---"
            )
        }
        catch{
            $message = Get-ExceptionMessage($_)
            Write-Log("CRITICAL: $message on $($Target.ComputerName)")
        }
    }

    function Stop-PowerShellSessions($runner)
    {
        foreach ($currentItem in $runner.TargetItems)
        {
            if ($null -ne $currentItem.Session)
            {
                Stop-PowerShellSession($currentItem)
            }
        }
    }

    function Get-Targets($Runner)
    {
        if ($Runner.FileType -eq "csv")
        {
            $targets = Import-Csv -Path $Runner.TargetName
        }
    
        if ($Runner.FileType -eq "computer")
        {
            $targets += [PSCustomObject]@{
                ComputerName = $Runner.TargetName
            }
        }

        $Runner.TotalTargets = $targets.Count -1 #index starts at 0
        $Runner.TargetItems = @()
        $itemIndex = 0
        foreach($item in $targets)
        {
            $Runner.TargetItems += [PSCustomObject]@{
            TargetId = New-Guid
            ComputerName = $targets[$itemIndex].ComputerName
            CachePath = $CachePath
            Service = $Service
            Session = $null
            ServiceStatus = $null
            CacheFound = $null
            CacheItems = $null
            CachedItemCount = $null
            TimeStarted = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            TimeEnded = $null
            StartFreeSpace = $null #new requirement
            EndFreeSpace = $null #new requirement
            Result = $null
            ResultDetails = $null
            JobNumber = 0
            }
            $itemIndex++
        }
    }

    Function Get-RunnerObject ($TargetName,$Service,$CachePath)
    {
        $batch = [PSCustomObject]@{
            BatchId = New-Guid
            Target = $Target
            TargetItems = $null
            Session = $null
            TimeStarted = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            TimeEnded = $null
            TotalTargets = 0
            CompletedTargets = 0
            FailedTargets = 0
            FileType = $null
            TargetName = $null
            ServiceName = $null
            CachePath = $null
            TimeElapsed = $null
        }
        $batch.CachePath = $CachePath
        $batch.ServiceName = $Service
        $batch.TargetName = $TargetName
        return $batch
    } 

    function Get-FileType ($Batch)
    {
            Write-Log -Message "Analyzing target type..."
            if (Test-Path -Path $Batch.TargetName)
            {
                if ($Target.ToLower().EndsWith(".csv")){
                    Write-Log -Message "Target determined as: .csv."
                    $Batch.FileType ="csv"
                }
                else {
                    throw (".csv file not provided. Please check the file type and try again.")
                }
            }
            else{
                Write-Log("Target determined as: single computer.")
                $Batch.FileType ="computer"
            }
    }

    function Get-TargetItems($Object)
    {
        $currentItem = $null
        foreach ($item in $Object.TargetItems)
        {
            try{
                Write-Log("---Starting work on: $($item.ComputerName)...---")
                $currentItem = $item
                $item.Session = New-PSSession -ComputerName $item.ComputerName -ErrorAction Stop
                Start-Actions $item
                Remove-PSSession $item.Session
                $Object.CompletedTargets++
                Write-Log("---Completed work on: $($item.ComputerName).---")
            }
            catch{
                $message = Get-ExceptionMessage($ex)
                $currentItem.ResultDetails = "CRITICAL: $message on $($currentItem.ComputerName)"
                $currentItem.Result = "Failed"
                Write-Log("CRITICAL: " + $_.Exception.Message)
                try {
                    if ($null -ne $currentItem.Session){
                        Remove-PSSession -Session $currentItem.Session -ErrorAction Stop
                    }
                    Write-Log("---Completed work on: $($currentItem.ComputerName).---")
                    $currentItem.Result = "Failed"
                    $Object.FailedTargets++
                    continue
                }
                catch {
                    Write-Log("CRITICAL: $($_.Exception.Message) on $($currentItem.Computername)")
                    Write-Log("---Completed work on: $($currentItem.ComputerName).---")
                    $currentItem.Result = "Failed"
                    $Object.FailedTargets++
                    continue
                }
            }
        }
    }

    function Remove-Files($Object){ 
        Write-Log("Deleting $($Object.CachedItemCount) items in MECM cache at: $($Object.CachePath) on $($Object.ComputerName)...")
        Invoke-Command -Session $Object.Session -Command {
            param($cache)
            Remove-Item -Path "$cache\*" -Recurse
        } -ArgumentList $Object.CachePath
        Write-Log("Completed deleting $($Object.CachedItemCount) items in MECM cache at: $($Object.CachePath) on $($Object.ComputerName).")
    }    

    function Get-FilesAtPath($Object)
    {
        Write-Log("Checking for items found in MECM cache at: $($Object.CachePath) on $($Object.ComputerName)...")
        $items = Invoke-Command -Session $Object.Session -Command {
            param($CacheFolderPath)
            Get-ChildItem -Path $CacheFolderPath -Recurse | Sort-Object -Property @{Expression={$_.FullName.Split('\').Count}; Descending=$true}
        } -ArgumentList $Object.CachePath
        if ($items.Length -eq 0)
        {
            Write-Log("No items found in MECM cache at: $($Object.CachePath) on $($Object.ComputerName)")
            $Object.CachedItemCount = 0
        }
        else{
            Write-Log("$($items.Length) items found in MECM cache at: $($Object.CachePath) on $($Object.ComputerName)")
            $Object.CachedItemCount =  $items.Length
        }
    }

    function Get-IsCacheFound($Object)
    {
        Write-Log("Checking if $($Object.CachePath) exists on: $($Object.ComputerName)...")
        $cacheStatus = Invoke-Command -Session $Object.Session -Command {
            param($Path)
            Test-Path -Path $Path
        } -ArgumentList $Object.CachePath
        
        if ($cacheStatus)
        {
            Write-Log("Found $($Object.CachePath) on: $($Object.ComputerName)...")
            $Object.CacheFound = $true
        }
        else{
            Write-Log("$($Object.CachePath) not found on $($Object.ComputerName)")
            $Object.CacheFound = $false
        } 
    }

    function Get-ServiceStatus($Object)
    {
        Write-Log("Checking Service Status: $($Object.Service) on: $($Object.ComputerName)")
        $Object.ServiceStatus = (Invoke-Command -Session $Object.Session -Command {
            param ($Name)
            Get-Service -Name $Name -ErrorAction SilentlyContinue
        } -ArgumentList $Object.Service).Status
        
        Write-Log("Service Status: $($Object.Service) is: $($Object.ServiceStatus) on: $($Object.ComputerName)")
    }

    function Stop-RemoteService($Object)
    {
        Write-Log("Attempting to stop service: $($Object.Service) on $($Object.ComputerName)...")
        Invoke-Command -Session $Object.Session -Command {
            param($Name)
                Stop-Service -Name $Name -Force
        } -ArgumentList $Object.Service
        Get-ServiceStatus -Object $Object

        while ($Object.ServiceStatus -ne "Stopped") #todo test
        {
            for ($i=1; $i -lt 3; $i++)
            {
                $message = "Service: $($Object.Service) was not stopped successfully on: $($Object.ComputerName), retrying..."
                Get-ServiceStatus -Object $Object
            }
            continue
        }

        if ($Object.ServiceStatus -eq "Stopped")
        {
            Write-Log("Service: $($Object.Service) was stopped successfully on: $($Object.ComputerName).")
        }
        else{
            $message = "Service: $($Object.Service) was not stopped successfully on: $($Object.ComputerName)."
            Write-Log($message)
            throw($message)
        }
    }

    function Start-RemoteService($Object)
    {
        Write-Log("Attempting to start service: $($Object.Service) on $($Object.ComputerName)...")
        Invoke-Command -Session $Object.Session -Command {
            param($Service)
            Start-Service -Name $Service
        } -ArgumentList $Object.Service
        Get-ServiceStatus -Object $Object
        if ($Object.ServiceStatus -eq "Running")
        {
            Write-Log("Service: $($Object.Service) was started successfully on: $($Object.ComputerName).")
        }
        else{
            Write-Log("Service: $($Object.Service) was not started successfully on: $($Object.ComputerName)")
            throw($_.Exception.Message, $Object)
        }
    }

    #todo: check for multiple disks
    function Get-FreeSpace($Object)
    {
        try {
            Write-Log("Checking free disk space on $($Object.ComputerName)")
            $freeSpace = (Invoke-Command -Session $Object.Session -Command {
            Get-PSDrive -PSProvider FileSystem | Select-Object Name, @{
                Name       = "FreeSpace"
                Expression = { [math]::round($_.Free / 1GB, 2) }
                }
            }).FreeSpace
            Write-Log("Found $freeSpace GB free disk space on $($Object.ComputerName)")
            if ($null -ne $Object.StartFreeSpace)
            {
                $Object.EndFreeSpace = $freeSpace
            }
            else{
                $Object.StartFreeSpace = $freeSpace
            }
        }
        catch {
            Write-Log("Free space check failed!")
            Write-Log("$($_.Exception.Message)")
            $Object.StartFreeSpace = 0
            $Object.EndFreeSpace = 0
            throw($_)
        }
    }
    
    function Get-Report($Object)
    {
        Write-Log("Generating report for BatchId:$($Object.BatchId)...")
        foreach ($item in $Object.TargetItems){
            if ($item.Result -eq "Success")
            {
                $Object.CompletedTargets++
            }
            else{
                $Object.FailedTargets++
            }
        }
        if ($Object.CompletedTargets -gt 0){
            Write-Log("Successfull targets: $($Object.CompletedTargets)")
            foreach ($item in $Object.TargetItems | Where-Object { $_.Result -eq "Success" }){
                Write-Log("Target: $($item.ComputerName) FreeSpaceAtStart: $($item.StartFreeSpace) GB FreeSpaceAtEnd: $($item.EndFreeSpace) GB")
            }
        }
        if ($Object.FailedTargets -gt 0){
            Write-Log("Failed Targets: $($Object.FailedTargets)")
            foreach ($item in $Object.TargetItems | Where-Object { $_.Result -eq "Failed" }){
                Write-Log("Target: $($item.ComputerName) `t FailureReason: $($item.ResultDetails)")
        }
    }   
        Write-Log("BatchId: $($Object.BatchId) Completed: $($Object.CompletedTargets) Failed: $($Object.FailedTargets)")
        Write-Log("Generating report for BatchId: $($Object.BatchId) completed.")
    }

    function Get-ElapsedTime($StartTime,$EndTime)
    {
        # Parse the timestamps into DateTime objects
        $format = "yyyy-MM-dd HH:mm:ss"
        $time1 = [datetime]::ParseExact($StartTime, $format, $null)
        $time2 = [datetime]::ParseExact($EndTime, $format, $null)

        # Calculate the difference
        $difference = $time2 - $time1

        # Extract hours, minutes, and seconds from the difference
        $hours = $difference.Hours
        $minutes = $difference.Minutes
        $seconds = $difference.Seconds

        # Format the difference as hours:minutes:seconds
        $difference_formatted = "{0:D2}:{1:D2}:{2:D2}" -f $hours, $minutes, $seconds
        return $difference_formatted
    }

    function Write-Log ($Message)
    {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = ("[$timestamp] $Message").ToString()
            Write-Output ($logMessage) #write output sends the output to the pipeline
    }

    function Get-ExceptionMessage($ex)
    {
                $truncatePast = 50
                $ex = $_.Exception.Message
                if ($ex.Length -gt $truncatePast)
                {
                    $message = $_.Exception.Message.Substring(0,$truncatePast)
                }
                else{
                    $message = $_.Exception.Message
                }
        return $message
    }
    #Beginning of main procedure
    Start-Routine -Target $Target -CachePath $CachePath -Service $Service #Start-Routine at end because other functions have to go before..I think :)..agous
}
