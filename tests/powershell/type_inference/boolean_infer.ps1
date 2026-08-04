# vybe-test: powershell/type_inference/boolean_infer
$x = $true
if ($x -is [bool]) { exit 0 }
exit 1
