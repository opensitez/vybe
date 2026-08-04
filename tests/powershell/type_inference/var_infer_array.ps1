# vybe-test: powershell/type_inference/var_infer_array
$x = 1,2,3
if ($x -is [object[]]) { exit 0 }
exit 1
