# vybe-test: powershell/background_jobs/start_job_alias
$job = Start-Job -ScriptBlock { 3 + 4 }
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result -ne 7) {
    Write-Host "FAIL: expected 7, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
