# vybe-test: powershell/collection_initializers/nested_array
$x = @(1, @(2,3))
if ($x[1][1] -eq 3) { exit 0 }
exit 1
