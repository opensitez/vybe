# vybe-test: powershell/hashtable_indexing/index_assignment
$x = @{ a = 5 }
$x['a'] = 6
if ($x['a'] -eq 6) { exit 0 }
exit 1
