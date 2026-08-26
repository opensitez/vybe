# vybe-test: powershell/collections_observable_collections/observable_collection_with_datetime_items
$oc = [System.Collections.ObjectModel.ObservableCollection[datetime]]::new()
$dt = [datetime]::UtcNow
$oc.Add($dt)
if ($oc[0] -ne $dt) { Write-Host "FAIL: DateTime item failed"; exit 1 }
Write-Host "PASS"; exit 0
