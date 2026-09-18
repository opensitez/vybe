# vybe-test: powershell/engine_history_management/history_buffer_cleared_to_empty
# After clearing the history buffer, calling Get-History evaluates cleanly to an empty collection
Clear-History
Add-History -InputObject ([pscustomobject]@{ CommandLine = "toBeCleared"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })

Clear-History
$entries = @(Get-History)

if ($entries.Count -ne 0) {
    Write-Host "FAIL: history entries remained after Clear-History, count: $($entries.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
