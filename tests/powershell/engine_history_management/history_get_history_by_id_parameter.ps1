# vybe-test: powershell/engine_history_management/history_get_history_by_id_parameter
# Get-History -Id retrieves the specific entry matching the given numeric ID
Clear-History
$first = Add-History -InputObject ([pscustomobject]@{ CommandLine = "TargetCommandForId"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now }) -Passthru

$retrieved = Get-History -Id $first.Id
Clear-History

if ($null -eq $retrieved) {
    Write-Host "FAIL: Get-History -Id returned `$null"
    exit 1
}

if ($retrieved.CommandLine -ne "TargetCommandForId") {
    Write-Host "FAIL: command line mismatch, got: '$($retrieved.CommandLine)'"
    exit 1
}

Write-Host "PASS"
exit 0
