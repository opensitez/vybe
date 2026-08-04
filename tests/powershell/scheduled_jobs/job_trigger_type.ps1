# vybe-test: powershell/scheduled_jobs/job_trigger_type
$trigger = New-JobTrigger -Daily -At (Get-Date).AddHours(1).TimeOfDay
if ($trigger.Type -ne 'Once' -and $trigger.Type -ne 'Daily') {
    Write-Host "FAIL: expected trigger type to be Created"
    exit 1
}
Write-Host "PASS"
exit 0
