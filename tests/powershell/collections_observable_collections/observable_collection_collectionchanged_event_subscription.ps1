# vybe-test: powershell/collections_observable_collections/observable_collection_collectionchanged_event_subscription
$oc = [System.Collections.ObjectModel.ObservableCollection[int]]::new()
$changed = $false
$action = { $script:changed = $true }
$oc.add_CollectionChanged($action)
$oc.Add(42)
$oc.remove_CollectionChanged($action)
if (-not $changed) { Write-Host "FAIL: CollectionChanged event failed"; exit 1 }
Write-Host "PASS"; exit 0
