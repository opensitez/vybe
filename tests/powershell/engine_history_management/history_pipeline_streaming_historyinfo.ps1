# vybe-test: powershell/engine_history_management/history_pipeline_streaming_historyinfo
# Add-History accepts HistoryInfo instances streamed directly through the pipeline
Clear-History
$entry = Add-History -InputObject ([pscustomobject]@{
    CommandLine = "PipedHistorySource"
    ExecutionStatus = "Completed"
    StartExecutionTime = [DateTime]::Now
    EndExecutionTime = [DateTime]::Now.AddMilliseconds(15)
}) -Passthru

Clear-History
$entry | Add-History

$restored = Get-History
Clear-History

if ($restored.Count -ne 1) {
    Write-Host "FAIL: expected 1 entry restored via pipeline, got $($restored.Count)"
    exit 1
}

if ($restored[0].CommandLine -ne "PipedHistorySource") {
    Write-Host "FAIL: command line mismatch, got: '$($restored[0].CommandLine)'"
    exit 1
}

Write-Host "PASS"
exit 0
