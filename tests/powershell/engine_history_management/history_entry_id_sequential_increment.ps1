# vybe-test: powershell/engine_history_management/history_entry_id_sequential_increment
# HistoryInfo .Id increases sequentially for each command added to history
Clear-History
$e1 = Add-History -InputObject ([pscustomobject]@{ CommandLine = "step 1"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now }) -Passthru
$e2 = Add-History -InputObject ([pscustomobject]@{ CommandLine = "step 2"; ExecutionStatus = "Completed"; StartExecutionTime = [DateTime]::Now; EndExecutionTime = [DateTime]::Now }) -Passthru
Clear-History

if ($e2.Id -le $e1.Id) {
    Write-Host "FAIL: history IDs did not increment, e1=$($e1.Id), e2=$($e2.Id)"
    exit 1
}

Write-Host "PASS"
exit 0
