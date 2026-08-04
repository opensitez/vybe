# vybe-test: powershell/hashtable_indexing/index_method
$x = @{ a=8 }
if ($x.Get_Item('a') -eq 8) { exit 0 }
exit 1
