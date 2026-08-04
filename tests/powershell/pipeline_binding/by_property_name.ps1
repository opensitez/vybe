# vybe-test: powershell/pipeline_binding/by_property_name
function Test { param([Parameter(ValueFromPipelineByPropertyName=$true)]$Name) process { $Name } }
[pscustomobject]@{ Name='x' } | Test | ForEach-Object { if ($_ -eq 'x') { exit 0 } }
exit 1
