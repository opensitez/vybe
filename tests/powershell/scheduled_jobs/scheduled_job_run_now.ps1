# vybe-test: powershell/scheduled_jobs/scheduled_job_run_now
$job = Start-ThreadJob -ScriptBlock { 10 + 20 }
$res = Receive-Job $job -Wait -AutoRemoveJob
if ($res -ne 30) {
    Write-Host "FAIL: ThreadJob failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
