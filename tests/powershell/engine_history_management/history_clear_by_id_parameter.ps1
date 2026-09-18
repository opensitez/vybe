# vybe-test: powershell/engine_history_management/history_clear_by_id_parameter
# Clear-History -Id removes only the specific entry matching the provided ID
Clear-History
$eA = Add-History -InputObject ([pscustomobject]@{ CommandLine = "KeepThisOne"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now }) -Passthru
$eB = Add-History -InputObject ([pscustomobject]@{ CommandLine = "DeleteThisOne"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now }) -Passthru

Clear-History -Id $eB.Id

$remaining = @(Get-History | ForEach-Object { $_.CommandLine })
Clear-History

if ($remaining -contains "DeleteThisOne") {
    Write-Host "FAIL: target command was not removed by Clear-History -Id"
    exit 1
}

if ($remaining -notcontains "KeepThisOne") {
    Write-Host "FAIL: non-target command was unexpectedly removed"
    exit 1
}

Write-Host "PASS"
exit 0
