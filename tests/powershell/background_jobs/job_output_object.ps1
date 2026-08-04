# vybe-test: powershell/background_jobs/job_output_object
$job = Start-Job -ScriptBlock { [PSCustomObject]@{ Value = 3 } }
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result.Value -ne 3) {
    Write-Host "FAIL: expected 3, got $($result.Value)"
    exit 1
}
Write-Host "PASS"
exit 0
