# vybe-test: powershell/engine_history_management/history_clear_all_entries
# Clear-History wipes all entries from the current session history buffer
Clear-History
Add-History -InputObject ([pscustomobject]@{ CommandLine = "temp1"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })
Add-History -InputObject ([pscustomobject]@{ CommandLine = "temp2"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })

Clear-History
$remaining = @(Get-History)

if ($remaining.Count -ne 0) {
    Write-Host "FAIL: history buffer was not cleared, remaining count: $($remaining.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
