# vybe-test: powershell/engine_history_management/history_clear_commandline_exact_match
# Clear-History -CommandLine exact string match removes only the exact matching entry
Clear-History
Add-History -InputObject ([pscustomobject]@{ CommandLine = "ExactTarget"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })
Add-History -InputObject ([pscustomobject]@{ CommandLine = "ExactTarget_Extension"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now })

Clear-History -CommandLine "ExactTarget"

$remaining = @(Get-History | ForEach-Object { $_.CommandLine })
Clear-History

if ($remaining -contains "ExactTarget") {
    Write-Host "FAIL: exact command was not removed"
    exit 1
}

if ($remaining -notcontains "ExactTarget_Extension") {
    Write-Host "FAIL: distinct command was improperly removed"
    exit 1
}

Write-Host "PASS"
exit 0
