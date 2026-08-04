# vybe-test: powershell/job_queues/queue_job_name
$job = Start-Job -Name QueueTest -ScriptBlock { 7 }
Wait-Job -Job $job
if ($job.Name -ne 'QueueTest') {
    Write-Host "FAIL: expected job name QueueTest, got $($job.Name)"
    exit 1
}
Write-Host "PASS"
exit 0
