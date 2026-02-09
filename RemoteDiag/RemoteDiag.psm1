Set-StrictMode -Version Latest

$script:RDDefaultLogPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'RemoteDiag.log'

function Write-RDLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',
        [string]$LogPath
    )

    $targetPath = if ([string]::IsNullOrWhiteSpace($LogPath)) { $script:RDDefaultLogPath } else { $LogPath }
    $parent = Split-Path -Path $targetPath -Parent
    if ($parent -and -not (Test-Path -Path $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fffK'), $Level, $Message
    Add-Content -Path $targetPath -Value $line -Encoding UTF8
}

function Test-RDHostReachability {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,
        [int]$TimeoutSeconds = 2,
        [string]$LogPath
    )

    Write-RDLog -LogPath $LogPath -Level 'DEBUG' -Message "Running ping pre-check for '$ComputerName' (TimeoutSeconds=$TimeoutSeconds)."
    try {
        $reachable = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -TimeoutSeconds $TimeoutSeconds -ErrorAction Stop
        if ($reachable) {
            Write-RDLog -LogPath $LogPath -Level 'DEBUG' -Message "Ping pre-check succeeded for '$ComputerName'."
            return $true
        }

        Write-RDLog -LogPath $LogPath -Level 'WARN' -Message "Ping pre-check did not receive a reply from '$ComputerName'."
        return $false
    }
    catch {
        Write-RDLog -LogPath $LogPath -Level 'WARN' -Message "Ping pre-check failed for '$ComputerName'. Error: $($_.Exception.Message)"
        return $false
    }
}

function Get-RDTransport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$LogPath,
        [ValidateSet('Auto', 'WinRM', 'RPC')]
        [string]$Preferred = 'Auto'
    )

    Write-RDLog -LogPath $LogPath -Level 'DEBUG' -Message "Transport selection started for '$ComputerName' (Preferred='$Preferred')."

    if ($Preferred -eq 'WinRM') {
        Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Transport forced to WinRM for '$ComputerName'."
        return 'WinRM'
    }
    if ($Preferred -eq 'RPC') {
        Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Transport forced to RPC for '$ComputerName'."
        return 'RPC'
    }

    try {
        $splat = @{ ComputerName = $ComputerName; ErrorAction = 'Stop' }
        if ($Credential) { $splat.Credential = $Credential }
        Test-WSMan @splat | Out-Null
        Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Transport probe succeeded via WinRM for '$ComputerName'."
        return 'WinRM'
    }
    catch {
        Write-RDLog -LogPath $LogPath -Level 'WARN' -Message "Transport probe failed for '$ComputerName'; falling back to RPC. Error: $($_.Exception.Message)"
        return 'RPC'
    }
}

function Invoke-RDRemote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,
        [hashtable]$ArgumentList,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$LogPath
    )

    $splat = @{ ComputerName = $ComputerName; ScriptBlock = $ScriptBlock; ErrorAction = 'Stop' }
    if ($null -ne $ArgumentList) {
        $ordered = $ArgumentList.GetEnumerator() | Sort-Object Name
        $splat.ArgumentList = @($ordered.Value)
    }
    if ($Credential) { $splat.Credential = $Credential }

    Write-RDLog -LogPath $LogPath -Level 'DEBUG' -Message "Invoking remote command against '$ComputerName'."
    Invoke-Command @splat
}

function Get-RDEventSummaryRPC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,
        [datetime]$StartTime,
        [string]$LogPath
    )

    $eventMap = @{
        AppCrashCount            = @{ LogName = 'Application'; Id = 1000 }
        KernelPowerCount         = @{ LogName = 'System'; Id = 41 }
        UnexpectedShutdownCount  = @{ LogName = 'System'; Id = 6008 }
        ResourceExhaustionCount  = @{ LogName = 'System'; Id = 2004 }
        BootDegradationEvents    = @{ LogName = 'Microsoft-Windows-Diagnostics-Performance/Operational'; Id = 100 }
    }

    $result = @{}
    foreach ($metric in $eventMap.Keys) {
        $filter = @{ LogName = $eventMap[$metric].LogName; Id = $eventMap[$metric].Id; StartTime = $StartTime }
        try {
            $count = (Get-WinEvent -ComputerName $ComputerName -FilterHashtable $filter -ErrorAction Stop | Measure-Object).Count
        }
        catch {
            $count = $null
        }
        $result[$metric] = $count
        Write-RDLog -LogPath $LogPath -Level 'DEBUG' -Message "RPC event query '$metric' for '$ComputerName' returned '$count'."
    }

    [pscustomobject]$result
}

function Get-RDHostSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]]$ComputerName,
        [int]$DaysBack = 7,
        [switch]$IncludeUpdates,
        [switch]$IncludeDefender,
        [switch]$IncludeWER,
        [switch]$Fast,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$LogPath,
        [ValidateSet('Auto', 'WinRM', 'RPC')]
        [string]$Transport = 'Auto'
    )

    begin {
        $startTime = (Get-Date).AddDays(-[math]::Abs($DaysBack))
        Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Host snapshot run started. Targets='$($ComputerName -join ',')'; DaysBack=$DaysBack; IncludeUpdates=$IncludeUpdates; IncludeDefender=$IncludeDefender; IncludeWER=$IncludeWER; Fast=$Fast; Transport=$Transport"
    }

    process {
        foreach ($target in $ComputerName) {
            $result = [ordered]@{
                ComputerName             = $target
                CollectionTime           = Get-Date
                TransportUsed            = $null
                Reachable                = $false
                FailureReason            = $null
                Domain                   = $null
                OSName                   = $null
                OSBuild                  = $null
                LastBootUpTime           = $null
                UptimeDays               = $null
                RAMTotalGB               = $null
                RAMFreeGB                = $null
                LowestDiskFreePct        = $null
                LowDiskFlag              = $false
                LowRAMFlag               = $false
                AppCrashCount            = 0
                KernelPowerCount         = 0
                UnexpectedShutdownCount  = 0
                ResourceExhaustionCount  = 0
                BootDegradationEvents    = 0
                HighCrashFlag            = $false
                DefenderLastQuickScan    = $null
                DefenderRecentDetections = $null
                WERReportCount           = $null
                DumpFileCount            = $null
                RecentUpdateCount        = $null
            }

            try {
                Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Collecting snapshot for '$target'."
                if (-not (Test-RDHostReachability -ComputerName $target -LogPath $LogPath)) {
                    $result.FailureReason = 'Ping pre-check failed; host did not respond before data collection started.'
                    Write-RDLog -LogPath $LogPath -Level 'WARN' -Message "Skipping snapshot for '$target' due to ping pre-check failure."
                    [pscustomobject]$result
                    continue
                }

                $selectedTransport = Get-RDTransport -ComputerName $target -Credential $Credential -LogPath $LogPath -Preferred $Transport
                $result.TransportUsed = $selectedTransport
                Write-RDLog -LogPath $LogPath -Level 'DEBUG' -Message "Selected transport '$selectedTransport' for '$target'."

                if ($selectedTransport -eq 'WinRM') {
                    $remote = Invoke-RDRemote -ComputerName $target -Credential $Credential -LogPath $LogPath -ArgumentList @{
                        IncludeUpdates = [bool]$IncludeUpdates
                        IncludeDefender = [bool]$IncludeDefender
                        IncludeWER = [bool]$IncludeWER
                        Fast = [bool]$Fast
                        StartTime = $startTime
                    } -ScriptBlock {
                        param($Fast, $IncludeDefender, $IncludeUpdates, $IncludeWER, $StartTime)

                        $os = Get-CimInstance Win32_OperatingSystem
                        $cs = Get-CimInstance Win32_ComputerSystem
                        $drives = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"

                        $eventMap = @{
                            AppCrashCount            = @{ LogName = 'Application'; Id = 1000 }
                            KernelPowerCount         = @{ LogName = 'System'; Id = 41 }
                            UnexpectedShutdownCount  = @{ LogName = 'System'; Id = 6008 }
                            ResourceExhaustionCount  = @{ LogName = 'System'; Id = 2004 }
                            BootDegradationEvents    = @{ LogName = 'Microsoft-Windows-Diagnostics-Performance/Operational'; Id = 100 }
                        }

                        $events = @{}
                        foreach ($metric in $eventMap.Keys) {
                            $filter = @{ LogName = $eventMap[$metric].LogName; Id = $eventMap[$metric].Id; StartTime = $StartTime }
                            try {
                                $events[$metric] = (Get-WinEvent -FilterHashtable $filter -ErrorAction Stop | Measure-Object).Count
                            }
                            catch {
                                $events[$metric] = $null
                            }
                        }

                        $defenderLast = $null
                        $defenderDetections = $null
                        if ($IncludeDefender -and -not $Fast) {
                            if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
                                try {
                                    $mp = Get-MpComputerStatus
                                    $defenderLast = $mp.QuickScanEndTime
                                }
                                catch { }
                            }

                            if (Get-WinEvent -ListLog 'Microsoft-Windows-Windows Defender/Operational' -ErrorAction SilentlyContinue) {
                                try {
                                    $defenderDetections = (Get-WinEvent -FilterHashtable @{
                                        LogName = 'Microsoft-Windows-Windows Defender/Operational'
                                        StartTime = $StartTime
                                        Id = 1116
                                    } -ErrorAction Stop | Measure-Object).Count
                                }
                                catch { }
                            }
                        }

                        $werReportCount = $null
                        $dumpCount = $null
                        if ($IncludeWER -and -not $Fast) {
                            try {
                                $werReportCount = (Get-ChildItem -Path 'C:\ProgramData\Microsoft\Windows\WER\ReportArchive' -Directory -ErrorAction Stop | Measure-Object).Count
                            }
                            catch { }

                            try {
                                $dumpCount = (Get-ChildItem -Path 'C:\Windows\Minidump' -Filter '*.dmp' -File -ErrorAction Stop | Measure-Object).Count
                            }
                            catch { }
                        }

                        $updateCount = $null
                        if ($IncludeUpdates -and -not $Fast) {
                            try {
                                $updateCount = (Get-HotFix | Where-Object { $_.InstalledOn -ge $StartTime } | Measure-Object).Count
                            }
                            catch { }
                        }

                        $lowestDiskPct = $null
                        if ($drives) {
                            $diskPcts = foreach ($d in $drives) {
                                if ($d.Size -gt 0) {
                                    [math]::Round(($d.FreeSpace / $d.Size) * 100, 2)
                                }
                            }
                            if ($diskPcts) {
                                $lowestDiskPct = ($diskPcts | Measure-Object -Minimum).Minimum
                            }
                        }

                        [pscustomobject]@{
                            Domain = $cs.Domain
                            OSName = $os.Caption
                            OSBuild = $os.BuildNumber
                            LastBootUpTime = $os.LastBootUpTime
                            UptimeDays = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 2)
                            RAMTotalGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
                            RAMFreeGB = [math]::Round($os.FreePhysicalMemory * 1KB / 1GB, 2)
                            LowestDiskFreePct = $lowestDiskPct
                            AppCrashCount = $events.AppCrashCount
                            KernelPowerCount = $events.KernelPowerCount
                            UnexpectedShutdownCount = $events.UnexpectedShutdownCount
                            ResourceExhaustionCount = $events.ResourceExhaustionCount
                            BootDegradationEvents = $events.BootDegradationEvents
                            DefenderLastQuickScan = $defenderLast
                            DefenderRecentDetections = $defenderDetections
                            WERReportCount = $werReportCount
                            DumpFileCount = $dumpCount
                            RecentUpdateCount = $updateCount
                        }
                    }

                    $result.Reachable = $true
                    foreach ($prop in $remote.PSObject.Properties.Name) {
                        $result[$prop] = $remote.$prop
                    }
                }
                else {
                    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $target -ErrorAction Stop
                    $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $target -ErrorAction Stop
                    $drives = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $target -Filter "DriveType=3" -ErrorAction SilentlyContinue
                    $events = Get-RDEventSummaryRPC -ComputerName $target -StartTime $startTime -LogPath $LogPath

                    $lowestDiskPct = $null
                    if ($drives) {
                        $diskPcts = foreach ($d in $drives) {
                            if ($d.Size -gt 0) {
                                [math]::Round(($d.FreeSpace / $d.Size) * 100, 2)
                            }
                        }
                        if ($diskPcts) {
                            $lowestDiskPct = ($diskPcts | Measure-Object -Minimum).Minimum
                        }
                    }

                    $result.Reachable = $true
                    $result.Domain = $cs.Domain
                    $result.OSName = $os.Caption
                    $result.OSBuild = $os.BuildNumber
                    $result.LastBootUpTime = $os.LastBootUpTime
                    $result.UptimeDays = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 2)
                    $result.RAMTotalGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
                    $result.RAMFreeGB = [math]::Round($os.FreePhysicalMemory * 1KB / 1GB, 2)
                    $result.LowestDiskFreePct = $lowestDiskPct
                    $result.AppCrashCount = $events.AppCrashCount
                    $result.KernelPowerCount = $events.KernelPowerCount
                    $result.UnexpectedShutdownCount = $events.UnexpectedShutdownCount
                    $result.ResourceExhaustionCount = $events.ResourceExhaustionCount
                    $result.BootDegradationEvents = $events.BootDegradationEvents
                }

                $result.LowDiskFlag = ($null -ne $result.LowestDiskFreePct -and $result.LowestDiskFreePct -lt 15)
                $result.LowRAMFlag = ($null -ne $result.RAMFreeGB -and $result.RAMFreeGB -lt 2)
                $totalCrashy = @($result.AppCrashCount, $result.KernelPowerCount, $result.UnexpectedShutdownCount) |
                    Where-Object { $null -ne $_ } |
                    Measure-Object -Sum
                $result.HighCrashFlag = ($totalCrashy.Sum -ge 5)
                Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Snapshot complete for '$target'. Reachable=$($result.Reachable); LowDiskFlag=$($result.LowDiskFlag); LowRAMFlag=$($result.LowRAMFlag); HighCrashFlag=$($result.HighCrashFlag); FailureReason=$($result.FailureReason)"
            }
            catch {
                $result.FailureReason = $_.Exception.Message
                Write-RDLog -LogPath $LogPath -Level 'ERROR' -Message "Snapshot failed for '$target'. Error: $($result.FailureReason)"
            }

            [pscustomobject]$result
        }
    }

    end {
        Write-RDLog -LogPath $LogPath -Level 'INFO' -Message 'Host snapshot run finished.'
    }
}

function Get-RDPerformanceSample {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ComputerName,
        [int]$DurationSeconds = 300,
        [int]$SampleIntervalSeconds = 5,
        [string[]]$Counters,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$LogPath
    )

    $defaultCounters = @(
        '\Processor(_Total)\% Processor Time',
        '\Memory\Available MBytes',
        '\PhysicalDisk(_Total)\Avg. Disk sec/Transfer',
        '\PhysicalDisk(_Total)\Avg. Disk Queue Length'
    )

    $counterList = if ($Counters) { $Counters } else { $defaultCounters }
    $sampleCount = [math]::Max([math]::Floor($DurationSeconds / $SampleIntervalSeconds), 1)
    Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Performance sample run started. Targets='$($ComputerName -join ',')'; DurationSeconds=$DurationSeconds; SampleIntervalSeconds=$SampleIntervalSeconds; CounterCount=$($counterList.Count)"

    foreach ($target in $ComputerName) {
        try {
            Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Collecting performance samples for '$target'."
            if (-not (Test-RDHostReachability -ComputerName $target -LogPath $LogPath)) {
                Write-RDLog -LogPath $LogPath -Level 'WARN' -Message "Skipping performance sample for '$target' due to ping pre-check failure."
                [pscustomobject]@{
                    ComputerName = $target
                    DurationSeconds = $DurationSeconds
                    SampleIntervalSeconds = $SampleIntervalSeconds
                    Counters = $counterList
                    Samples = @()
                    Summary = @()
                    FailureReason = 'Ping pre-check failed; host did not respond before performance collection started.'
                }
                continue
            }

            $splat = @{
                Counter = $counterList
                ComputerName = $target
                SampleInterval = $SampleIntervalSeconds
                MaxSamples = $sampleCount
                ErrorAction = 'Stop'
            }
            if ($Credential) { $splat.Credential = $Credential }

            $raw = Get-Counter @splat

            $samples = foreach ($entry in $raw.CounterSamples) {
                [pscustomobject]@{
                    ComputerName = $target
                    TimeStamp = $entry.TimeStamp
                    CounterPath = $entry.Path
                    Value = [math]::Round($entry.CookedValue, 4)
                }
            }

            $grouped = $samples | Group-Object CounterPath
            $summary = foreach ($g in $grouped) {
                [pscustomobject]@{
                    ComputerName = $target
                    CounterPath = $g.Name
                    Average = [math]::Round((($g.Group.Value | Measure-Object -Average).Average), 4)
                    Minimum = [math]::Round((($g.Group.Value | Measure-Object -Minimum).Minimum), 4)
                    Maximum = [math]::Round((($g.Group.Value | Measure-Object -Maximum).Maximum), 4)
                }
            }

            [pscustomobject]@{
                ComputerName = $target
                DurationSeconds = $DurationSeconds
                SampleIntervalSeconds = $SampleIntervalSeconds
                Counters = $counterList
                Samples = $samples
                Summary = $summary
                FailureReason = $null
            }
            Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Performance sample complete for '$target'. Samples=$($samples.Count); SummaryRows=$($summary.Count)"
        }
        catch {
            Write-RDLog -LogPath $LogPath -Level 'ERROR' -Message "Performance sample failed for '$target'. Error: $($_.Exception.Message)"
            [pscustomobject]@{
                ComputerName = $target
                DurationSeconds = $DurationSeconds
                SampleIntervalSeconds = $SampleIntervalSeconds
                Counters = $counterList
                Samples = @()
                Summary = @()
                FailureReason = $_.Exception.Message
            }
        }
    }

    Write-RDLog -LogPath $LogPath -Level 'INFO' -Message 'Performance sample run finished.'
}

function Export-RDReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]$InputObject,
        [Parameter(Mandatory)]
        [string]$Path,
        [string]$LogPath,
        [ValidateSet('Json', 'Csv', 'Html')]
        [string]$Format = 'Json',
        [switch]$IncludeRawEvents
    )

    begin {
        $buffer = @()
    }

    process {
        $buffer += $InputObject
    }

    end {
        Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Export started. Format='$Format'; Path='$Path'; ItemCount=$($buffer.Count)"
        $parent = Split-Path -Path $Path -Parent
        if ($parent -and -not (Test-Path $parent)) {
            New-Item -Path $parent -ItemType Directory -Force | Out-Null
        }

        switch ($Format) {
            'Json' {
                $buffer | ConvertTo-Json -Depth 8 | Set-Content -Path $Path -Encoding UTF8
            }
            'Csv' {
                $flattened = $buffer | ForEach-Object {
                    $_ | Select-Object * -ExcludeProperty Samples, Summary
                }
                $flattened | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
            }
            'Html' {
                $buffer | ConvertTo-Html -Title 'RemoteDiag Report' | Set-Content -Path $Path -Encoding UTF8
            }
        }

        Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Export finished. Format='$Format'; Path='$Path'"
        Get-Item -Path $Path
    }
}

function Compare-RDHostSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Reference,
        [Parameter(Mandatory)]
        [pscustomobject]$Difference,
        [string]$LogPath
    )

    Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Snapshot comparison started. Reference='$($Reference.ComputerName)'; Difference='$($Difference.ComputerName)'"

    $checks = @(
        @{ Metric = 'OSBuild'; Reference = $Reference.OSBuild; Difference = $Difference.OSBuild; Direction = 'Mismatch' }
        @{ Metric = 'UptimeDays'; Reference = $Reference.UptimeDays; Difference = $Difference.UptimeDays; Direction = 'Delta' }
        @{ Metric = 'LowestDiskFreePct'; Reference = $Reference.LowestDiskFreePct; Difference = $Difference.LowestDiskFreePct; Direction = 'LowerIsBad' }
        @{ Metric = 'RAMFreeGB'; Reference = $Reference.RAMFreeGB; Difference = $Difference.RAMFreeGB; Direction = 'LowerIsBad' }
        @{ Metric = 'AppCrashCount'; Reference = $Reference.AppCrashCount; Difference = $Difference.AppCrashCount; Direction = 'HigherIsBad' }
        @{ Metric = 'KernelPowerCount'; Reference = $Reference.KernelPowerCount; Difference = $Difference.KernelPowerCount; Direction = 'HigherIsBad' }
        @{ Metric = 'UnexpectedShutdownCount'; Reference = $Reference.UnexpectedShutdownCount; Difference = $Difference.UnexpectedShutdownCount; Direction = 'HigherIsBad' }
        @{ Metric = 'ResourceExhaustionCount'; Reference = $Reference.ResourceExhaustionCount; Difference = $Difference.ResourceExhaustionCount; Direction = 'HigherIsBad' }
        @{ Metric = 'BootDegradationEvents'; Reference = $Reference.BootDegradationEvents; Difference = $Difference.BootDegradationEvents; Direction = 'HigherIsBad' }
    )

    $rows = foreach ($check in $checks) {
        $delta = $null
        if (($check.Reference -is [ValueType]) -and ($check.Difference -is [ValueType])) {
            $delta = [double]$check.Difference - [double]$check.Reference
        }

        $severity = 'Info'
        switch ($check.Direction) {
            'Mismatch' { if ($check.Reference -ne $check.Difference) { $severity = 'Warning' } }
            'HigherIsBad' { if ($delta -gt 0) { $severity = 'Warning' } }
            'LowerIsBad' { if ($delta -lt 0) { $severity = 'Warning' } }
            'Delta' { if ([math]::Abs($delta) -gt 2) { $severity = 'Info' } }
        }

        [pscustomobject]@{
            Metric = $check.Metric
            Reference = $check.Reference
            Difference = $check.Difference
            Delta = $delta
            Severity = $severity
        }
    }

    $likely = @()
    if ($Difference.LowestDiskFreePct -lt 15) { $likely += 'Low free disk likely contributes to slowness/crashes.' }
    if ($Difference.ResourceExhaustionCount -gt $Reference.ResourceExhaustionCount) { $likely += 'Resource exhaustion events are elevated.' }
    if (($Difference.AppCrashCount + $Difference.KernelPowerCount + $Difference.UnexpectedShutdownCount) -
        ($Reference.AppCrashCount + $Reference.KernelPowerCount + $Reference.UnexpectedShutdownCount) -gt 3) {
        $likely += 'Crash-related events are significantly higher than reference.'
    }
    if ($Difference.BootDegradationEvents -gt $Reference.BootDegradationEvents) { $likely += 'Boot degradation events are elevated.' }

    $comparison = [pscustomobject]@{
        ReferenceComputer = $Reference.ComputerName
        DifferenceComputer = $Difference.ComputerName
        Comparison = $rows
        LikelyContributors = $likely
    }

    Write-RDLog -LogPath $LogPath -Level 'INFO' -Message "Snapshot comparison finished. WarningCount=$(($rows | Where-Object Severity -eq 'Warning').Count); LikelyContributors=$($likely.Count)"
    $comparison
}

Export-ModuleMember -Function Get-RDHostSnapshot, Get-RDPerformanceSample, Export-RDReport, Compare-RDHostSnapshot
