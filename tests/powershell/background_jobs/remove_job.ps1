# vybe-test: powershell/background_jobs/remove_job
$job = Start-Job -ScriptBlock { Start-Sleep -Milliseconds 1; 1 }
Wait-Job -Job $job
Remove-Job -Job $job
if (Get-Job -Id $job.Id -ErrorAction SilentlyContinue) {
    Write-Host "FAIL: expected job removed"
    exit 1
}
Write-Host "PASS"
exit 0
