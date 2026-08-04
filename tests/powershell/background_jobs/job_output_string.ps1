# vybe-test: powershell/background_jobs/job_output_string
$job = Start-Job -ScriptBlock { "abc" }
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result -ne 'abc') {
    Write-Host "FAIL: expected abc, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
