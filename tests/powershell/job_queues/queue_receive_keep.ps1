# vybe-test: powershell/job_queues/queue_receive_keep
$job = Start-Job -ScriptBlock { 12 }
Wait-Job -Job $job
$first = Receive-Job -Job $job -Keep
$second = Receive-Job -Job $job -Keep
if ($first -ne 12 -or $second -ne 12) {
    Write-Host "FAIL: expected two identical receives"
    exit 1
}
Write-Host "PASS"
exit 0
