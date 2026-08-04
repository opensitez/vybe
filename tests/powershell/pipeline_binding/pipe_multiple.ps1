# vybe-test: powershell/pipeline_binding/pipe_multiple
function Inc { process { $_ + 1 } }
if ((1,2 | Inc | Measure-Object).Count -eq 2) { exit 0 }
exit 1
