# vybe-test: powershell/pipeline_binding/value_by_value
function Test { param([Parameter(ValueFromPipeline=$true)]$InputValue) process { $_ + $InputValue } }
if ((1 | Test -InputValue 2) -eq 3) { exit 0 }
exit 1
