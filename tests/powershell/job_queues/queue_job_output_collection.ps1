# vybe-test: powershell/job_queues/queue_job_output_collection
$job = Start-Job -ScriptBlock { 1; 2; 3 }
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result -ne 3) {
    Write-Host "FAIL: expected three output objects, got $($result.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
