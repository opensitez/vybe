# vybe-test: powershell/collections_readonly_collections/readonly_collection_in_foreach_loop
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3))
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[int]]::new($list)
$sum = 0
foreach ($item in $roc) { $sum += $item }
if ($sum -ne 6) { Write-Host "FAIL: foreach on ReadOnlyCollection failed"; exit 1 }
Write-Host "PASS"; exit 0
