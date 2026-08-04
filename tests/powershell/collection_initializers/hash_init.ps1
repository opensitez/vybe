# vybe-test: powershell/collection_initializers/hash_init
$x = @{ a = 1; b = 2 }
if ($x['a'] -eq 1 -and $x['b'] -eq 2) { exit 0 }
exit 1
