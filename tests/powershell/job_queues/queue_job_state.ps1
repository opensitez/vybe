# vybe-test: powershell/job_queues/queue_job_state
$job = Start-Job -ScriptBlock { Start-Sleep -Seconds 1; 1 }
if ($job.State -notin 'Running','Completed','Suspended') {
    Write-Host "FAIL: expected valid job state, got $($job.State)"
    exit 1
}
Wait-Job -Job $job
Write-Host "PASS"
exit 0
