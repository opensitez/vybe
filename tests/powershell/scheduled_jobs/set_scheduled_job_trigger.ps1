# vybe-test: powershell/scheduled_jobs/set_scheduled_job_trigger
$job = Register-ScheduledJob -Name TestTrigger -ScriptBlock { 3 + 3 } -RunNow
Set-ScheduledJob -Name TestTrigger -Trigger (New-JobTrigger -Once -At (Get-Date).AddMinutes(1)) | Out-Null
$job = Get-ScheduledJob -Name TestTrigger
if (-not $job) {
    Write-Host "FAIL: expected scheduled job to still exist"
    exit 1
}
Unregister-ScheduledJob -Name TestTrigger -Force -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
