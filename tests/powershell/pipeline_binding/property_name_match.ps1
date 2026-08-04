# vybe-test: powershell/pipeline_binding/property_name_match
function Test { param([Parameter(ValueFromPipelineByPropertyName=$true)]$Name) process { $Name } }
[pscustomobject]@{ Name='hello' } | Test | ForEach-Object { if ($_ -eq 'hello') { exit 0 } }
exit 1
