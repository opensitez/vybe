# vybe-test: powershell/type_inference/pipeline_infer
$x = 1..3 | ForEach-Object { $_ }
if ($x -is [object[]]) { exit 0 }
exit 1
