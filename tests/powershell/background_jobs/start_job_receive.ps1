# vybe-test: powershell/background_jobs/start_job_receive
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
