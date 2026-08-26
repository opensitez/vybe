# vybe-test: powershell/collections_readonly_collections/readonly_collection_nested_in_psobject
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1, 2, 3))
$obj = [pscustomobject]@{ ReadOnlyItems = [System.Collections.ObjectModel.ReadOnlyCollection[int]]::new($list) }
if ($obj.ReadOnlyItems.Count -ne 3) { Write-Host "FAIL: Nested ReadOnlyCollection in PSCustomObject failed"; exit 1 }
Write-Host "PASS"; exit 0
