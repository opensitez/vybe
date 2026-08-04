# vybe-test: powershell/job_queues/queue_job_verbose
$job = Start-Job -ScriptBlock { Write-Verbose 'v' } -Verbose
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result -ne $null) {
    Write-Host "FAIL: expected no direct output from verbose"
    exit 1
}
Write-Host "PASS"
exit 0
