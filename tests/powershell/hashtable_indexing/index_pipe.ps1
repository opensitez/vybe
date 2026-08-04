# vybe-test: powershell/hashtable_indexing/index_pipe
$x = @{ a = 1 }
if ((1 | ForEach-Object { $x['a'] }) -eq 1) { exit 0 }
exit 1
