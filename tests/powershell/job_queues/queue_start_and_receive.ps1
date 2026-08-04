# vybe-test: powershell/job_queues/queue_start_and_receive
$job = Start-Job -ScriptBlock { 10 }
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result -ne 10) {
    Write-Host "FAIL: expected 10, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
