# vybe-test: powershell/hashtable_indexing/index_nested
$x = @{ outer = @{ inner = 3 } }
if ($x['outer']['inner'] -eq 3) { exit 0 }
exit 1
