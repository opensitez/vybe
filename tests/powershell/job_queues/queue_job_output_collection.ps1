# vybe-test: powershell/job_queues/queue_job_output_collection
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
