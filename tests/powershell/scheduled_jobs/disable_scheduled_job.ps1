# vybe-test: powershell/scheduled_jobs/disable_scheduled_job
$job = Register-ScheduledJob -Name DisableJob -ScriptBlock { 'x' }
Disable-ScheduledJob -Name DisableJob
$details = Get-ScheduledJob -Name DisableJob
if ($details.Enabled) {
    Write-Host "FAIL: expected scheduled job disabled"
    exit 1
}
Unregister-ScheduledJob -Name DisableJob -Force -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
