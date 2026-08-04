# vybe-test: powershell/collection_initializers/hash_entry
$x = @{ Name='abc'; Value=5 }
if ($x.Name -eq 'abc') { exit 0 }
exit 1
