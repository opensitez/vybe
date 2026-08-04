# vybe-test: powershell/hashtable_indexing/index_literal
$x = @{ a = 1 }
if ($x['a'] -eq 1) { exit 0 }
exit 1
