Add-Type -AssemblyName PresentationFramework
Import-Module "$PSScriptRoot/../RemoteDiag/RemoteDiag.psd1" -Force

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="RemoteDiag" Height="600" Width="1000">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,8">
            <TextBlock Text="Targets (comma-separated):" VerticalAlignment="Center"/>
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

        <DataGrid x:Name="ResultsGrid" Grid.Row="2" AutoGenerateColumns="True" IsReadOnly="True"/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

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

$snapshotResults = @()
$logPath = Join-Path $pwd 'RemoteDiag-Activity.log'

$runButton.Add_Click({
    $statusText.Text = 'Running...'
    Add-Content -Path $logPath -Value "$(Get-Date -Format 's') [INFO] GUI snapshot requested. Targets=$($targetsBox.Text)"
    $window.Dispatcher.Invoke([action]{}, 'Render')

    $targets = $targetsBox.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $days = 7
    [void][int]::TryParse($daysBackBox.Text, [ref]$days)

    if (-not $targets) {
        $statusText.Text = 'Enter at least one target.'
        return
    }

    try {
        $snapshotResults = Get-RDHostSnapshot -ComputerName $targets -DaysBack $days -IncludeWER:$includeWER.IsChecked -IncludeDefender:$includeDefender.IsChecked -IncludeUpdates:$includeUpdates.IsChecked -LogPath $logPath
        $resultsGrid.ItemsSource = $snapshotResults
        $statusText.Text = "Completed: $($snapshotResults.Count) targets"
        Add-Content -Path $logPath -Value "$(Get-Date -Format 's') [INFO] GUI snapshot completed. ResultCount=$($snapshotResults.Count)"
    }
    catch {
        $statusText.Text = "Error: $($_.Exception.Message)"
        Add-Content -Path $logPath -Value "$(Get-Date -Format 's') [ERROR] GUI snapshot failed. Error=$($_.Exception.Message)"
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
