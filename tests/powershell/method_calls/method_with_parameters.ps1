# vybe-test: powershell/method_calls/method_with_parameters
if ('hello'.Replace('h','j') -eq 'jello') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
