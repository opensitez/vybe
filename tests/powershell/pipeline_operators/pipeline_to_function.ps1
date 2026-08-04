# vybe-test: powershell/pipeline_operators/pipeline_to_function
function Test-Func { process { $_ } }
if ((1,2 | Test-Func) -join ',' -eq '1,2') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
