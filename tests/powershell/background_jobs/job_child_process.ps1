# vybe-test: powershell/background_jobs/job_child_process
$job = Start-Job -ScriptBlock { Get-Process -Id $PID }
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result.Id -ne $result.Id) {
    Write-Host "FAIL: expected process object"
    exit 1
}
Write-Host "PASS"
exit 0
