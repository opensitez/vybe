# vybe-test: powershell/type_inference/object_infer
$x = New-Object PSObject
if ($x -is [object]) { exit 0 }
exit 1
