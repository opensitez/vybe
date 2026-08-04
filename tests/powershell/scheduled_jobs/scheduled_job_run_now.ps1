# vybe-test: powershell/scheduled_jobs/scheduled_job_run_now
Register-ScheduledJob -Name RunNowJob -ScriptBlock { 'run' } -RunNow
$job = Get-ScheduledJob -Name RunNowJob
if (-not $job) {
    Write-Host "FAIL: expected scheduled job to be registered"
    exit 1
}
Unregister-ScheduledJob -Name RunNowJob -Force -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
