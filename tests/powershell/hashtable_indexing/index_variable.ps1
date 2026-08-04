# vybe-test: powershell/hashtable_indexing/index_variable
$key = 'a'
$x = @{ a = 2 }
if ($x[$key] -eq 2) { exit 0 }
exit 1
