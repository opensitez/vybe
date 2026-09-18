# vybe-test: powershell/engine_history_management/history_clear_newest_with_count
# Clear-History -Newest -Count N removes the N most recently executed entries
Clear-History
Add-History -InputObject ([pscustomobject]@{ CommandLine = "first_kept"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })
Add-History -InputObject ([pscustomobject]@{ CommandLine = "second_deleted"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })

Clear-History -Newest -Count 1

$remaining = @(Get-History | ForEach-Object { $_.CommandLine })
Clear-History

if ($remaining -contains "second_deleted") {
    Write-Host "FAIL: newest entry was not removed by Clear-History -Newest -Count 1"
    exit 1
}

if ($remaining -notcontains "first_kept") {
    Write-Host "FAIL: older entry was unexpectedly removed"
    exit 1
}

Write-Host "PASS"
exit 0
