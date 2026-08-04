# vybe-test: powershell/hashtable_indexing/index_dynamic
$key = 'a'
$x = @{ a = 4 }
if ($x[$key] -eq 4) { exit 0 }
exit 1
