# vybe-test: powershell/scheduled_jobs/start_scheduled_job
Register-ScheduledJob -Name TestStart -ScriptBlock { 2 + 2 } -RunNow
Start-ScheduledJob -Name TestStart | Out-Null
$jobs = Get-Job -State Running,Completed | Where-Object { $_.Name -eq 'TestStart' }
if (-not $jobs) {
    Write-Host "FAIL: expected scheduled job to start"
    Remove-Job -Name TestStart -ErrorAction SilentlyContinue
    exit 1
}
Get-ScheduledJob -Name TestStart | Unregister-ScheduledJob -Force -ErrorAction SilentlyContinue
Write-Host "PASS"
exit 0
