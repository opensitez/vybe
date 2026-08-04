# vybe-test: powershell/collection_initializers/array_comma
$x = 1,2,3
if ($x[2] -eq 3) { exit 0 }
exit 1
