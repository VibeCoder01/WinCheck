Set-StrictMode -Version Latest

function Get-RDTransport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential,
        [ValidateSet('Auto', 'WinRM', 'RPC')]
        [string]$Preferred = 'Auto'
    )

    if ($Preferred -eq 'WinRM') { return 'WinRM' }
    if ($Preferred -eq 'RPC') { return 'RPC' }

    try {
        $splat = @{ ComputerName = $ComputerName; ErrorAction = 'Stop' }
        if ($Credential) { $splat.Credential = $Credential }
        Test-WSMan @splat | Out-Null
        return 'WinRM'
    }
    catch {
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
        [System.Management.Automation.PSCredential]$Credential
    )

    $splat = @{ ComputerName = $ComputerName; ScriptBlock = $ScriptBlock; ErrorAction = 'Stop' }
    if ($null -ne $ArgumentList) {
        $ordered = $ArgumentList.GetEnumerator() | Sort-Object Name
        $splat.ArgumentList = @($ordered.Value)
    }
    if ($Credential) { $splat.Credential = $Credential }

    Invoke-Command @splat
}

function Get-RDEventSummaryRPC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,
        [datetime]$StartTime
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
        [ValidateSet('Auto', 'WinRM', 'RPC')]
        [string]$Transport = 'Auto'
    )

    begin {
        $startTime = (Get-Date).AddDays(-[math]::Abs($DaysBack))
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
                $selectedTransport = Get-RDTransport -ComputerName $target -Credential $Credential -Preferred $Transport
                $result.TransportUsed = $selectedTransport

                if ($selectedTransport -eq 'WinRM') {
                    $remote = Invoke-RDRemote -ComputerName $target -Credential $Credential -ArgumentList @{
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
                    $events = Get-RDEventSummaryRPC -ComputerName $target -StartTime $startTime

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
            }
            catch {
                $result.FailureReason = $_.Exception.Message
            }

            [pscustomobject]$result
        }
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
        [System.Management.Automation.PSCredential]$Credential
    )

    $defaultCounters = @(
        '\Processor(_Total)\% Processor Time',
        '\Memory\Available MBytes',
        '\PhysicalDisk(_Total)\Avg. Disk sec/Transfer',
        '\PhysicalDisk(_Total)\Avg. Disk Queue Length'
    )

    $counterList = if ($Counters) { $Counters } else { $defaultCounters }
    $sampleCount = [math]::Max([math]::Floor($DurationSeconds / $SampleIntervalSeconds), 1)

    foreach ($target in $ComputerName) {
        try {
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
        }
        catch {
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
}

function Export-RDReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]$InputObject,
        [Parameter(Mandatory)]
        [string]$Path,
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

        Get-Item -Path $Path
    }
}

function Compare-RDHostSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Reference,
        [Parameter(Mandatory)]
        [pscustomobject]$Difference
    )

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

    [pscustomobject]@{
        ReferenceComputer = $Reference.ComputerName
        DifferenceComputer = $Difference.ComputerName
        Comparison = $rows
        LikelyContributors = $likely
    }
}

Export-ModuleMember -Function Get-RDHostSnapshot, Get-RDPerformanceSample, Export-RDReport, Compare-RDHostSnapshot
