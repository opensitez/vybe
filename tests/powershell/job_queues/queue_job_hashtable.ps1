# vybe-test: powershell/job_queues/queue_job_hashtable
$job = Start-Job -ScriptBlock { [hashtable]@{ A = 1 } }
Wait-Job -Job $job
$result = Receive-Job -Job $job
if ($result['A'] -ne 1) {
    Write-Host "FAIL: expected hashtable value 1"
    exit 1
}
Write-Host "PASS"
exit 0
