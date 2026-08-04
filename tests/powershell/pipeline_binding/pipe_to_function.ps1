# vybe-test: powershell/pipeline_binding/pipe_to_function
function Add-One { param($x) process { $_ + $x } }
if ((1 | Add-One -x 2) -eq 3) { exit 0 }
exit 1
