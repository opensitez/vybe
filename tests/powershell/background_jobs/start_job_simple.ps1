# vybe-test: powershell/background_jobs/start_job_simple
$job = Start-Job -ScriptBlock { 1 + 1 }
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result -ne 2) {
    Write-Host "FAIL: expected 2, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
