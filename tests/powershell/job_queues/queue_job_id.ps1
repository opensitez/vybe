# vybe-test: powershell/job_queues/queue_job_id
$job = Start-Job -ScriptBlock { 9 }
if ($job.Id -lt 1) {
    Write-Host "FAIL: expected job id >= 1"
    exit 1
}
Wait-Job -Job $job
Write-Host "PASS"
exit 0
