# vybe-test: powershell/hashtable_indexing/index_missing
$x = @{ a = 1 }
if ($x['missing'] -eq $null) { exit 0 }
exit 1
