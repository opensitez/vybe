# vybe-test: powershell/background_jobs/job_receive_twice
$job = Start-ThreadJob -ScriptBlock { 42 }
$res = Receive-Job $job -Wait -AutoRemoveJob
if ($res -ne 42) {
    Write-Host "FAIL: ThreadJob receive failed"
    exit 1
}
Write-Host "PASS"
exit 0
