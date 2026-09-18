# vybe-test: powershell/background_jobs/job_receive_twice
$job = Start-Job -ScriptBlock { 42 }
Wait-Job -Job $job
$result1 = Receive-Job -Job $job
$result2 = Receive-Job -Job $job -Keep
if ($result1 -ne 42 -or $result2 -ne 42) {
    Write-Host "FAIL: expected two receives both 42"
    exit 1
}
Write-Host "PASS"
exit 0
