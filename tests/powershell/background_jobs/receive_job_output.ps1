# vybe-test: powershell/background_jobs/receive_job_output
$job = Start-Job -ScriptBlock { Write-Output 'hello' }
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result -ne 'hello') {
    Write-Host "FAIL: expected hello, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
