# vybe-test: powershell/collections_observable_collections/observable_collection_count_property
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new()
for ($i = 0; $i -lt 40; $i++) { $oc.Add($i) }
if ($oc.Count -ne 40) { Write-Host "FAIL: Count property failed"; exit 1 }
Write-Host "PASS"; exit 0
