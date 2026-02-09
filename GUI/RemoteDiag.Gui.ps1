$startupLogPath = Join-Path $PSScriptRoot 'RemoteDiag.Gui.startup.log'

if (-not $IsWindows) {
    throw 'RemoteDiag GUI only runs on Windows hosts with WPF support.'
}

trap {
    $errorMessage = $_.Exception.Message
    Add-Content -Path $startupLogPath -Value "$(Get-Date -Format 's') [FATAL] GUI startup failed. Error=$errorMessage"
    if ([type]::GetType('System.Windows.MessageBox, PresentationFramework')) {
        [System.Windows.MessageBox]::Show("RemoteDiag GUI failed to start: $errorMessage`nSee log: $startupLogPath", 'RemoteDiag Startup Error') | Out-Null
    }
    Write-Error "RemoteDiag GUI failed to start: $errorMessage (see $startupLogPath)"
    break
}

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $currentProcessPath = (Get-Process -Id $PID).Path
    if (-not $currentProcessPath -or -not (Test-Path -Path $currentProcessPath -PathType Leaf)) {
        $currentProcessPath = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
    }
    if (-not $currentProcessPath -or -not (Test-Path -Path $currentProcessPath -PathType Leaf)) {
        $currentProcessPath = (Get-Command powershell.exe -ErrorAction SilentlyContinue).Source
    }

    if (-not $currentProcessPath) {
        throw 'Unable to determine a PowerShell executable to relaunch GUI in STA mode.'
    }

    $argumentList = @('-NoProfile', '-STA', '-File', $PSCommandPath) + $args

    Write-Warning 'RemoteDiag GUI requires an STA runspace. Relaunching in STA mode...'
    Add-Content -Path $startupLogPath -Value "$(Get-Date -Format 's') [INFO] Relaunching GUI in STA mode. Host=$currentProcessPath Args=$($argumentList -join ' ')"
    Start-Process -FilePath $currentProcessPath -ArgumentList $argumentList -WorkingDirectory $PSScriptRoot | Out-Null
    exit
}

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Import-Module "$PSScriptRoot/../RemoteDiag/RemoteDiag.psd1" -Force

# Avoid WPF GPU/desktop-heap pressure issues in constrained or remote sessions.
[System.Windows.Media.RenderOptions]::ProcessRenderMode = [System.Windows.Interop.RenderMode]::SoftwareOnly

if (-not [System.Windows.Application]::Current) {
    $null = [System.Windows.Application]::new()
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="RemoteDiag" Height="600" Width="1000" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,8">
            <TextBlock Text="Targets (CSV, range, or file):" VerticalAlignment="Center"/>
            <TextBox x:Name="TargetsBox" Width="400" Margin="8,0,0,0"/>
            <TextBlock Text="Days back:" VerticalAlignment="Center" Margin="12,0,0,0"/>
            <TextBox x:Name="DaysBackBox" Width="50" Text="7" Margin="6,0,0,0"/>
            <CheckBox x:Name="IncludeWER" Content="WER" Margin="12,0,0,0"/>
            <CheckBox x:Name="IncludeDefender" Content="Defender" Margin="8,0,0,0"/>
            <CheckBox x:Name="IncludeUpdates" Content="Updates" Margin="8,0,0,0"/>
        </StackPanel>

        <StackPanel Orientation="Horizontal" Grid.Row="1" Margin="0,0,0,8">
            <Button x:Name="RunButton" Content="Run Snapshot" Width="120"/>
            <Button x:Name="ExportJsonButton" Content="Export JSON" Width="100" Margin="8,0,0,0"/>
            <Button x:Name="ExportCsvButton" Content="Export CSV" Width="100" Margin="8,0,0,0"/>
            <TextBlock x:Name="StatusText" Margin="16,4,0,0" Text="Ready"/>
        </StackPanel>

        <StackPanel Orientation="Horizontal" Grid.Row="2" Margin="0,0,0,8">
            <TextBlock Text="Progress:" VerticalAlignment="Center"/>
            <ProgressBar x:Name="RunProgressBar" Width="320" Height="16" Margin="8,0,0,0" Minimum="0" Maximum="1" Value="0"/>
            <TextBlock x:Name="ProgressText" Margin="8,0,0,0" VerticalAlignment="Center" Text="0/0"/>
        </StackPanel>

        <DataGrid x:Name="ResultsGrid" Grid.Row="3" AutoGenerateColumns="True" IsReadOnly="True"/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$window.Topmost = $true
$window.Add_ContentRendered({ $window.Topmost = $false; $window.Activate() })

$targetsBox = $window.FindName('TargetsBox')
$daysBackBox = $window.FindName('DaysBackBox')
$includeWER = $window.FindName('IncludeWER')
$includeDefender = $window.FindName('IncludeDefender')
$includeUpdates = $window.FindName('IncludeUpdates')
$runButton = $window.FindName('RunButton')
$exportJsonButton = $window.FindName('ExportJsonButton')
$exportCsvButton = $window.FindName('ExportCsvButton')
$resultsGrid = $window.FindName('ResultsGrid')
$statusText = $window.FindName('StatusText')
$runProgressBar = $window.FindName('RunProgressBar')
$progressText = $window.FindName('ProgressText')

$snapshotResults = @()
$liveResults = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$resultsGrid.ItemsSource = $liveResults
$logPath = Join-Path $pwd 'RemoteDiag-Activity.log'
Add-Content -Path $startupLogPath -Value "$(Get-Date -Format 's') [INFO] GUI startup initialized. Script=$PSCommandPath"
$modulePath = Join-Path $PSScriptRoot '../RemoteDiag/RemoteDiag.psd1'
$runTimer = [System.Windows.Threading.DispatcherTimer]::new()
$runTimer.Interval = [TimeSpan]::FromMilliseconds(250)
$activeRunspace = $null
$activeInvocation = $null
$outputBuffer = $null
$outputDataAddedHandler = $null
$totalTargets = 0
$completedTargets = 0


function Expand-TargetToken {
    param(
        [Parameter(Mandatory)]
        [string]$Token
    )

    $trimmed = $Token.Trim()
    if (-not $trimmed) {
        return @()
    }

    $rangeMatch = [regex]::Match($trimmed, '^(?<start>\S+)\s*(?<sep>\.\.|-)\s*(?<end>\S+)$')
    if (-not $rangeMatch.Success) {
        return @($trimmed)
    }

    $startName = $rangeMatch.Groups['start'].Value
    $endName = $rangeMatch.Groups['end'].Value
    $startMatch = [regex]::Match($startName, '^(?<prefix>.*?)(?<number>\d+)$')
    $endMatch = [regex]::Match($endName, '^(?<prefix>.*?)(?<number>\d+)$')
    if (-not $startMatch.Success -or -not $endMatch.Success) {
        throw "Invalid range '$trimmed'. Both range endpoints must end with digits."
    }

    if ($startMatch.Groups['prefix'].Value -ne $endMatch.Groups['prefix'].Value) {
        throw "Invalid range '$trimmed'. Range endpoints must share the same prefix."
    }

    $startNumber = [int]$startMatch.Groups['number'].Value
    $endNumber = [int]$endMatch.Groups['number'].Value
    if ($endNumber -lt $startNumber) {
        throw "Invalid range '$trimmed'. Ending number must be greater than or equal to starting number."
    }

    $prefix = $startMatch.Groups['prefix'].Value
    $width = [Math]::Max($startMatch.Groups['number'].Value.Length, $endMatch.Groups['number'].Value.Length)

    $expanded = foreach ($number in $startNumber..$endNumber) {
        '{0}{1}' -f $prefix, $number.ToString("D$width")
    }

    return @($expanded)
}

function Get-TargetsFromFile {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        throw "Target file '$Path' does not exist."
    }

    $targets = foreach ($line in (Get-Content -Path $Path -ErrorAction Stop)) {
        $trimmed = $line.TrimStart()
        if (-not $trimmed) { continue }
        if ($trimmed[0] -notmatch '[A-Za-z0-9]') { continue }

        $machineName = ($trimmed -split '\s+', 2)[0]
        if ($machineName) {
            $machineName
        }
    }

    return @($targets)
}

function Resolve-ComputerTargets {
    param(
        [Parameter(Mandatory)]
        [string]$RawInput
    )

    $allTargets = New-Object System.Collections.Generic.List[string]
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $items = $RawInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    foreach ($item in $items) {
        $isFile = ($item -match '\.[A-Za-z0-9]+$') -and (Test-Path -Path $item -PathType Leaf)
        $resolved = if ($isFile) {
            Get-TargetsFromFile -Path $item
        }
        else {
            Expand-TargetToken -Token $item
        }

        foreach ($target in $resolved) {
            if ($seen.Add($target)) {
                $null = $allTargets.Add($target)
            }
        }
    }

    return @($allTargets)
}

function Set-RunState {
    param([bool]$IsRunning)

    $runButton.IsEnabled = -not $IsRunning
    $exportJsonButton.IsEnabled = -not $IsRunning
    $exportCsvButton.IsEnabled = -not $IsRunning
}

function Update-ProgressDisplay {
    if ($totalTargets -le 0) {
        $runProgressBar.Minimum = 0
        $runProgressBar.Maximum = 1
        $runProgressBar.Value = 0
        $progressText.Text = '0/0'
        return
    }

    $runProgressBar.Minimum = 0
    $runProgressBar.Maximum = $totalTargets
    $runProgressBar.Value = [Math]::Min($completedTargets, $totalTargets)
    $progressText.Text = "$completedTargets/$totalTargets"
}

$runTimer.Add_Tick({
    if (-not $activeInvocation -or -not $activeInvocation.IsCompleted) {
        return
    }

    $runTimer.Stop()

    try {
        $null = $activeRunspace.EndInvoke($activeInvocation)

        while ($outputBuffer -and $liveResults.Count -lt $outputBuffer.Count) {
            $liveResults.Add($outputBuffer[$liveResults.Count])
        }

        $snapshotResults = @($liveResults)
        $completedTargets = $snapshotResults.Count
        Update-ProgressDisplay
        $statusText.Text = "Completed: $($snapshotResults.Count) targets"
        Add-Content -Path $logPath -Value "$(Get-Date -Format 's') [INFO] GUI snapshot completed. ResultCount=$($snapshotResults.Count)"
    }
    catch {
        $statusText.Text = "Error: $($_.Exception.Message)"
        Add-Content -Path $logPath -Value "$(Get-Date -Format 's') [ERROR] GUI snapshot failed. Error=$($_.Exception.Message)"
    }
    finally {
        if ($outputBuffer -and $outputDataAddedHandler) {
            $outputBuffer.remove_DataAdded($outputDataAddedHandler)
        }
        if ($activeRunspace) {
            $activeRunspace.Dispose()
        }
        $activeRunspace = $null
        $activeInvocation = $null
        $outputBuffer = $null
        $outputDataAddedHandler = $null
        Set-RunState -IsRunning $false
    }
})

$runButton.Add_Click({
    if ($activeInvocation -and -not $activeInvocation.IsCompleted) {
        return
    }

    $statusText.Text = 'Running...'
    Add-Content -Path $logPath -Value "$(Get-Date -Format 's') [INFO] GUI snapshot requested. Targets=$($targetsBox.Text)"
    $window.Dispatcher.Invoke([action]{}, 'Render')

    try {
        $targets = Resolve-ComputerTargets -RawInput $targetsBox.Text
    }
    catch {
        $statusText.Text = "Target parsing error: $($_.Exception.Message)"
        Add-Content -Path $logPath -Value "$(Get-Date -Format 's') [ERROR] GUI target parse failed. Error=$($_.Exception.Message)"
        return
    }
    $days = 7
    [void][int]::TryParse($daysBackBox.Text, [ref]$days)

    if (-not $targets) {
        $statusText.Text = 'Enter at least one target.'
        return
    }

    $liveResults.Clear()
    $snapshotResults = @()
    $totalTargets = $targets.Count
    $completedTargets = 0
    Update-ProgressDisplay

    Set-RunState -IsRunning $true

    $activeRunspace = [powershell]::Create()
    $null = $activeRunspace.AddScript({
        param($ModulePath, $Targets, $Days, $IncludeWer, $IncludeDefender, $IncludeUpdates, $LogPath)

        Import-Module $ModulePath -Force
        Get-RDHostSnapshot -ComputerName $Targets -DaysBack $Days -IncludeWER:$IncludeWer -IncludeDefender:$IncludeDefender -IncludeUpdates:$IncludeUpdates -LogPath $LogPath
    })
    $null = $activeRunspace.AddArgument($modulePath)
    $null = $activeRunspace.AddArgument($targets)
    $null = $activeRunspace.AddArgument($days)
    $null = $activeRunspace.AddArgument([bool]$includeWER.IsChecked)
    $null = $activeRunspace.AddArgument([bool]$includeDefender.IsChecked)
    $null = $activeRunspace.AddArgument([bool]$includeUpdates.IsChecked)
    $null = $activeRunspace.AddArgument($logPath)

    $inputBuffer = [System.Management.Automation.PSDataCollection[psobject]]::new()
    $inputBuffer.Complete()
    $outputBuffer = [System.Management.Automation.PSDataCollection[psobject]]::new()

    $outputDataAddedHandler = {
        param($sender, $eventArgs)

        $nextResult = $sender[$eventArgs.Index]
        $window.Dispatcher.BeginInvoke([action]{
            $liveResults.Add($nextResult)
            $completedTargets = $liveResults.Count
            Update-ProgressDisplay
            $statusText.Text = "Running... $completedTargets/$totalTargets"
        }) | Out-Null
    }
    $outputBuffer.add_DataAdded($outputDataAddedHandler)

    $activeInvocation = $activeRunspace.BeginInvoke($inputBuffer, $outputBuffer)
    $runTimer.Start()
})

$window.Add_Closing({
    $runTimer.Stop()
    if ($activeRunspace) {
        try {
            $activeRunspace.Stop()
        }
        catch { }
        $activeRunspace.Dispose()
    }
})

$exportJsonButton.Add_Click({
    if (-not $snapshotResults) {
        $statusText.Text = 'Nothing to export.'
        return
    }
    $path = Join-Path $pwd 'RemoteDiag-Snapshot.json'
    $snapshotResults | Export-RDReport -Path $path -Format Json -LogPath $logPath | Out-Null
    $statusText.Text = "Exported: $path"
    Add-Content -Path $logPath -Value "$(Get-Date -Format 's') [INFO] GUI exported JSON report. Path=$path"
})

$exportCsvButton.Add_Click({
    if (-not $snapshotResults) {
        $statusText.Text = 'Nothing to export.'
        return
    }
    $path = Join-Path $pwd 'RemoteDiag-Snapshot.csv'
    $snapshotResults | Export-RDReport -Path $path -Format Csv -LogPath $logPath | Out-Null
    $statusText.Text = "Exported: $path"
    Add-Content -Path $logPath -Value "$(Get-Date -Format 's') [INFO] GUI exported CSV report. Path=$path"
})

$null = $window.ShowDialog()
