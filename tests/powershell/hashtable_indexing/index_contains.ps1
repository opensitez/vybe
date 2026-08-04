# vybe-test: powershell/hashtable_indexing/index_contains
$x = @{ a = 9 }
if ($x.ContainsKey('a')) { exit 0 }
exit 1
