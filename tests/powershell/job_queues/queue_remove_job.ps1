# vybe-test: powershell/job_queues/queue_remove_job
$job = Start-Job -ScriptBlock { 13 }
Wait-Job -Job $job
Remove-Job -Job $job -Force
if (Get-Job -Id $job.Id -ErrorAction SilentlyContinue) {
    Write-Host "FAIL: expected removed job"
    exit 1
}
Write-Host "PASS"
exit 0
