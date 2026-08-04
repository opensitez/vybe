# vybe-test: powershell/scheduled_jobs/enable_scheduled_job
$job = Register-ScheduledJob -Name EnableJob -ScriptBlock { 'x' }
Disable-ScheduledJob -Name EnableJob
Enable-ScheduledJob -Name EnableJob
$details = Get-ScheduledJob -Name EnableJob
if (-not $details.Enabled) {
    Write-Host "FAIL: expected scheduled job enabled"
    exit 1
}
Unregister-ScheduledJob -Name EnableJob -Force -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
