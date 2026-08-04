# vybe-test: powershell/scheduled_jobs/unregister_scheduled_job
Register-ScheduledJob -Name RemoveJob -ScriptBlock { 5 }
Unregister-ScheduledJob -Name RemoveJob -ErrorAction SilentlyContinue
$job = Get-ScheduledJob -Name RemoveJob -ErrorAction SilentlyContinue
if ($job) {
    Write-Host "FAIL: expected no scheduled job after unregister"
    exit 1
}
Write-Host "PASS"
exit 0
