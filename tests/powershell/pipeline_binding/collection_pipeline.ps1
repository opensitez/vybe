# vybe-test: powershell/pipeline_binding/collection_pipeline
function Test { param($x) process { $_ * $x } }
if ((1,2 | Test -x 2 | Measure-Object).Count -eq 2) { exit 0 }
exit 1
