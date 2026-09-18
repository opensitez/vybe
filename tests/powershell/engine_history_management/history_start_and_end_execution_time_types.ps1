# vybe-test: powershell/engine_history_management/history_start_and_end_execution_time_types
# HistoryInfo exposes StartExecutionTime and EndExecutionTime properties of type [DateTime]
Clear-History
$entry = Add-History -InputObject ([pscustomobject]@{
    CommandLine = "TimeTypeCheck"
    ExecutionStatus = "Completed"
    StartExecutionTime = [DateTime]::Now
    EndExecutionTime = [DateTime]::Now.AddMilliseconds(5)
}) -Passthru
Clear-History

if (-not ($entry.StartExecutionTime -is [DateTime])) {
    Write-Host "FAIL: StartExecutionTime is not a DateTime, got: $($entry.StartExecutionTime.GetType().FullName)"
    exit 1
}

if (-not ($entry.EndExecutionTime -is [DateTime])) {
    Write-Host "FAIL: EndExecutionTime is not a DateTime, got: $($entry.EndExecutionTime.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
