# vybe-test: powershell/pipeline_binding/explicit_parameter
function Test { param($x) process { $_ + $x } }
if ((1 | Test -x 2) -eq 3) { exit 0 }
exit 1
