# vybe-test: powershell/pipeline_binding/complex_input
function Test { param([Parameter(ValueFromPipeline=$true)]$Number) process { $_ + $Number } }
if ((2 | Test -Number 3) -eq 5) { exit 0 }
exit 1
