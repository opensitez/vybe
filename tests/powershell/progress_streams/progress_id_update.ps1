# vybe-test: powershell/progress_streams/progress_id_update
Write-Progress -Id 2 -Activity 'Updating' -Status 'Half' -PercentComplete 50
Write-Progress -Id 2 -Activity 'Updating' -Status 'Almost' -PercentComplete 90
Write-Host 'PASS'
exit 0
