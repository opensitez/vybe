# vybe-test: powershell/engine_history_management/history_entry_duration_timespan
# HistoryInfo .Duration is computed as the TimeSpan difference between EndExecutionTime and StartExecutionTime
Clear-History
$start = [DateTime]::Parse("2026-01-01 12:00:00")
$end   = [DateTime]::Parse("2026-01-01 12:00:05")

$entry = Add-History -InputObject ([pscustomobject]@{
    CommandLine = "LongRunningJob"
    ExecutionStatus = "Completed"
    StartExecutionTime = $start
    EndExecutionTime = $end
}) -Passthru
Clear-History

if (-not ($entry.Duration -is [TimeSpan])) {
    Write-Host "FAIL: Duration is not a TimeSpan, got: $($entry.Duration.GetType().FullName)"
    exit 1
}

if ($entry.Duration.TotalSeconds -ne 5) {
    Write-Host "FAIL: expected Duration of 5 seconds, got: $($entry.Duration.TotalSeconds)"
    exit 1
}

Write-Host "PASS"
exit 0
