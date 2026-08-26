# vybe-test: powershell/collections_observable_collections/observable_collection_pipeline_filter
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new([int[]]@(1..10))
$gt5 = @($oc | Where-Object { $_ -gt 5 })
if ($gt5.Length -ne 5) { Write-Host "FAIL: Pipeline filter failed"; exit 1 }
Write-Host "PASS"; exit 0
