# vybe-test: powershell/collections_readonly_collections/readonly_collection_to_array_via_linq_or_loop
$list = [System.Collections.Generic.List[int]]::new([int[]]@(10, 20, 30))
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[int]]::new($list)
$arr = [int[]]::new($roc.Count)
for ($i = 0; $i -lt $roc.Count; $i++) { $arr[$i] = $roc[$i] }
if ($arr.Length -ne 3 -or $arr[0] -ne 10) { Write-Host "FAIL: Array export failed"; exit 1 }
Write-Host "PASS"; exit 0
