# vybe-test: powershell/parameter_binding/pipeline_parameter
function Test-Func { param($x); process { $_ } }
if ((1 | Test-Func) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
