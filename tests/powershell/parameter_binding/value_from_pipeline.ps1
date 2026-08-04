# vybe-test: powershell/parameter_binding/value_from_pipeline
function Test-Func { param([Parameter(ValueFromPipeline=$true)]$x); process { $_ } }
if ((1 | Test-Func) -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
