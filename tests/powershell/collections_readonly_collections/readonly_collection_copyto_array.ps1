# vybe-test: powershell/collections_readonly_collections/readonly_collection_copyto_array
$list = [System.Collections.Generic.List[int]]::new([int[]]@(5, 10))
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[int]]::new($list)
[int[]]$arr = [int[]]::new(2)
$roc.CopyTo($arr, 0)
if ($arr[0] -ne 5 -or $arr[1] -ne 10) { Write-Host "FAIL: CopyTo failed"; exit 1 }
Write-Host "PASS"; exit 0
