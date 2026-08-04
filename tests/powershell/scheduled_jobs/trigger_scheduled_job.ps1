# vybe-test: powershell/scheduled_jobs/trigger_scheduled_job
Register-ScheduledJob -Name TriggerJob -ScriptBlock { Write-Output 'ok' } -RunNow
$trigger = New-JobTrigger -Once -At (Get-Date).AddSeconds(1)
Set-ScheduledJob -Name TriggerJob -Trigger $trigger | Out-Null
if (-not (Get-ScheduledJob -Name TriggerJob)) {
    Write-Host "FAIL: expected scheduled job to exist"
    exit 1
}
Unregister-ScheduledJob -Name TriggerJob -Force -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
