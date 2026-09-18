# vybe-test: powershell/engine_history_management/history_get_history_count_parameter
# Get-History -Count limits the number of most recent history entries returned
Clear-History
1..5 | ForEach-Object {
    Add-History -InputObject ([pscustomobject]@{ CommandLine = "cmd $_"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })
}

$limited = @(Get-History -Count 2)
Clear-History

if ($limited.Count -ne 2) {
    Write-Host "FAIL: expected 2 entries with -Count 2, got $($limited.Count)"
    exit 1
}

if ($limited[1].CommandLine -ne "cmd 5") {
    Write-Host "FAIL: expected latest command 'cmd 5', got: '$($limited[1].CommandLine)'"
    exit 1
}

Write-Host "PASS"
exit 0
