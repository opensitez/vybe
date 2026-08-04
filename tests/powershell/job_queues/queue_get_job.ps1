# vybe-test: powershell/job_queues/queue_get_job
Start-Job -Name QueryJob -ScriptBlock { 8 } | Out-Null
$job = Get-Job -Name QueryJob
if ($job.Name -ne 'QueryJob') {
    Write-Host "FAIL: expected to retrieve QueryJob"
    exit 1
}
Get-Job -Name QueryJob | Remove-Job -Force
Write-Host "PASS"
exit 0
