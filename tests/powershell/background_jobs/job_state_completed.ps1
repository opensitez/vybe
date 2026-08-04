# vybe-test: powershell/background_jobs/job_state_completed
$job = Start-Job -ScriptBlock { "done" }
Wait-Job -Job $job
if ($job.State -ne 'Completed') {
    Write-Host "FAIL: expected Completed, got $($job.State)"
    exit 1
}
Write-Host "PASS"
exit 0
