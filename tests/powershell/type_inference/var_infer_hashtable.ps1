# vybe-test: powershell/type_inference/var_infer_hashtable
$x = @{ a = 1 }
if ($x -is [hashtable]) { exit 0 }
exit 1
