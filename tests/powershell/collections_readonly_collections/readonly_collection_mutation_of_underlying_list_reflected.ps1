# vybe-test: powershell/collections_readonly_collections/readonly_collection_mutation_of_underlying_list_reflected
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2))
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[int]]::new($list)
$list.Add(3)
if ($roc.Count -ne 3 -or $roc[2] -ne 3) { Write-Host "FAIL: Underlying mutation reflection failed"; exit 1 }
Write-Host "PASS"; exit 0
