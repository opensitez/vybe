# vybe-test: powershell/parameter_binding/value_from_pipeline_by_property_name
function Test-Func { param([Parameter(ValueFromPipelineByPropertyName=$true)]$Name); process { $_.Name } }
$obj = [pscustomobject]@{ Name='PASS' }
if (($obj | Test-Func) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
