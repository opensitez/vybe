# vybe-test: powershell/progress_streams/progress_status
Write-Progress -Activity 'Status' -Status 'Running' -PercentComplete 30
Write-Host 'PASS'
exit 0
