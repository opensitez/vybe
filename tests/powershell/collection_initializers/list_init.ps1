# vybe-test: powershell/collection_initializers/list_init
$x = [system.collections.generic.list[int]]@(1,2)
if ($x.Count -eq 2) { exit 0 }
exit 1
