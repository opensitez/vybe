# vybe-test: powershell/type_inference/mixed_infer
$x = 1 + 2
if ($x -is [int]) { exit 0 }
exit 1
