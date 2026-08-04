# vybe-test: powershell/type_inference/string_concat_infer
$x = 'a' + 'b'
if ($x -is [string]) { exit 0 }
exit 1
