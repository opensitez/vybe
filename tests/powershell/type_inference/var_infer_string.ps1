# vybe-test: powershell/type_inference/var_infer_string
$x = 'hello'
if ($x -is [string]) { exit 0 }
exit 1
