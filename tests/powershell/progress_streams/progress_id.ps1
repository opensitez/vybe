# vybe-test: powershell/progress_streams/progress_id
Write-Progress -Id 1 -Activity 'A' -Status 'S' -PercentComplete 10
Write-Host 'PASS'
exit 0
