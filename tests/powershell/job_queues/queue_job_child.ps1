# vybe-test: powershell/job_queues/queue_job_child
$job = Start-Job -ScriptBlock { Get-Process -Id $PID }
Wait-Job -Job $job
$result = Receive-Job -Job $job
if (-not $result.Id) {
    Write-Host "FAIL: expected process object"
    exit 1
}
Write-Host "PASS"
exit 0
