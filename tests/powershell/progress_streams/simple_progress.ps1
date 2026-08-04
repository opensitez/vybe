# vybe-test: powershell/progress_streams/simple_progress
Write-Progress -Activity 'Test' -Status 'Running' -PercentComplete 25
Write-Host 'PASS'
exit 0
