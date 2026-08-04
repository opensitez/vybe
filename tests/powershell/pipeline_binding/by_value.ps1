# vybe-test: powershell/pipeline_binding/by_value
function Test { param($x) process { $_ + 1 } }
if ((1..2 | Test | Measure-Object).Count -eq 2) { exit 0 }
exit 1
