# vybe-test: powershell/collection_initializers/empty_array
$x = @()
if ($x.Count -eq 0) { exit 0 }
exit 1
