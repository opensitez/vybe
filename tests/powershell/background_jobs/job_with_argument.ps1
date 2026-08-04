# vybe-test: powershell/background_jobs/job_with_argument
$script = { param($x) $x * 2 }
$job = Start-Job -ScriptBlock $script -ArgumentList 5
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result -ne 10) {
    Write-Host "FAIL: expected 10, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
