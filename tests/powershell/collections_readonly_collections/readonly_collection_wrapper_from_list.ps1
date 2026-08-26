# vybe-test: powershell/collections_readonly_collections/readonly_collection_wrapper_from_list
$list = [System.Collections.Generic.List[int]]::new([int[]]@(10, 20, 30))
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[int]]::new($list)
if ($roc.Count -ne 3 -or $roc[1] -ne 20) { Write-Host "FAIL: ReadOnlyCollection wrapper failed"; exit 1 }
Write-Host "PASS"; exit 0
