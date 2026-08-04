# vybe-test: powershell/progress_streams/progress_multiple
Write-Progress -Activity 'One' -Status 'Step1' -PercentComplete 10
Write-Progress -Activity 'Two' -Status 'Step2' -PercentComplete 20
Write-Host 'PASS'
exit 0
