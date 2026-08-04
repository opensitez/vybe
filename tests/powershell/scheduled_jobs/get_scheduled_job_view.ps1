# vybe-test: powershell/scheduled_jobs/get_scheduled_job_view
$job = Register-ScheduledJob -Name ViewJob -ScriptBlock { 'x' }
$found = Get-ScheduledJob -Name ViewJob
if ($found.Name -ne 'ViewJob') {
    Write-Host "FAIL: expected scheduled job name ViewJob"
    exit 1
}
Unregister-ScheduledJob -Name ViewJob -Force -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
