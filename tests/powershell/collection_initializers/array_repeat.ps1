# vybe-test: powershell/collection_initializers/array_repeat
$x = ,1
if ($x.Count -eq 1) { exit 0 }
exit 1
