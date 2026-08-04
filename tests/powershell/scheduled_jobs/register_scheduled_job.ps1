# vybe-test: powershell/scheduled_jobs/register_scheduled_job
Register-ScheduledJob -Name TestJob -ScriptBlock { 1 + 1 } -RunNow
$job = Get-ScheduledJob -Name TestJob -ErrorAction SilentlyContinue
if (-not $job) {
    Write-Host "FAIL: expected scheduled job to exist"
    exit 1
}
Unregister-ScheduledJob -Name TestJob -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
