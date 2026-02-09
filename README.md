# WinCheck RemoteDiag

`RemoteDiag` is a PowerShell 7 module and GUI wrapper for repeatable remote triage of Windows endpoints.

## Included commands

- `Get-RDHostSnapshot`: Collect baseline host state + high-value event counts.
- `Get-RDPerformanceSample`: Capture short perf-counter time series with summary stats.
- `Export-RDReport`: Export output to JSON, CSV, or HTML.
- `Compare-RDHostSnapshot`: Diff a “good” reference vs “bad” host with simple likely-cause heuristics.

## MVP coverage

This implementation includes:

- OS build, uptime, RAM/disk health metrics.
- Event counts for IDs 1000, 41, 6008, 2004, Diagnostics-Performance 100.
- Optional Defender / WER / updates collection.
- Transport selection (`WinRM` preferred, `RPC` fallback).
- Per-target fault isolation (`FailureReason` per machine, no whole-run stop).
- JSON/CSV/HTML export.
- Basic WPF GUI shell (`GUI/RemoteDiag.Gui.ps1`) for running snapshots and exporting results.

## Quick start

```powershell
Import-Module ./RemoteDiag/RemoteDiag.psd1 -Force

$snapshots = Get-RDHostSnapshot -ComputerName PC001,PC002 -DaysBack 7
$snapshots | Export-RDReport -Path ./out/snapshot.json -Format Json
$snapshots | Format-Table ComputerName,OSBuild,UptimeDays,LowestDiskFreePct,AppCrashCount,FailureReason

$comparison = Compare-RDHostSnapshot -Reference $snapshots[0] -Difference $snapshots[1]
$comparison.Comparison | Format-Table
$comparison.LikelyContributors
```

To launch GUI (on Windows PowerShell host with WPF support):

```powershell
pwsh ./GUI/RemoteDiag.Gui.ps1
```
