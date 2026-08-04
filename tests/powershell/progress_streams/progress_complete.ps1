# vybe-test: powershell/progress_streams/progress_complete
Write-Progress -Activity 'End' -Status 'Complete' -PercentComplete 100
Write-Host 'PASS'
exit 0
