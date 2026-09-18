# vybe-test: powershell/engine_history_management/history_get_history_all_entries
# Get-History without parameters returns all recorded entries in the current session
Clear-History
Add-History -InputObject ([pscustomobject]@{ CommandLine = "command 1"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })
Add-History -InputObject ([pscustomobject]@{ CommandLine = "command 2"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })

$allEntries = @(Get-History)
Clear-History

if ($allEntries.Count -ne 2) {
    Write-Host "FAIL: expected 2 history entries, got $($allEntries.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
