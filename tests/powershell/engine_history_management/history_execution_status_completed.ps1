# vybe-test: powershell/engine_history_management/history_execution_status_completed
# HistoryInfo .ExecutionStatus property reflects execution completion
Clear-History
$entry = Add-History -InputObject ([pscustomobject]@{
    CommandLine = "StatusVerification"
    ExecutionStatus = "Completed"
    StartExecutionTime = [DateTime]::Now
    EndExecutionTime = [DateTime]::Now
}) -Passthru
Clear-History

if ($entry.ExecutionStatus.ToString() -ne "Completed") {
    Write-Host "FAIL: ExecutionStatus mismatch, expected 'Completed', got: '$($entry.ExecutionStatus)'"
    exit 1
}

Write-Host "PASS"
exit 0
