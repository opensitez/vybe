# vybe-test: powershell/hashtable_indexing/index_hashtable_property
$x = @{ a = 7 }
if ($x.a -eq 7) { exit 0 }
exit 1
