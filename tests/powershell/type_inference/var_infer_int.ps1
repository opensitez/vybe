# vybe-test: powershell/type_inference/var_infer_int
$x = 10
if ($x -is [int]) { exit 0 }
exit 1
