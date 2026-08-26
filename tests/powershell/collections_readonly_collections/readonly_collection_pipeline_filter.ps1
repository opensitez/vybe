# vybe-test: powershell/collections_readonly_collections/readonly_collection_pipeline_filter
$list = [System.Collections.Generic.List[int]]::new([int[]]@(1..10))
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[int]]::new($list)
$evens = @($roc | Where-Object { $_ % 2 -eq 0 })
if ($evens.Length -ne 5) { Write-Host "FAIL: Pipeline filter failed"; exit 1 }
Write-Host "PASS"; exit 0
