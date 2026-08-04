# vybe-test: powershell/collection_initializers/array_init
$x = 1,2,3
if ($x.Count -eq 3) { exit 0 }
exit 1
