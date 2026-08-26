# vybe-test: powershell/collections_observable_collections/observable_collection_with_guid_items
$oc = [System.Collections.ObjectModel.ObservableCollection[guid]]::new()
$g = [guid]::NewGuid()
$oc.Add($g)
if ($oc[0] -ne $g) { Write-Host "FAIL: Guid item failed"; exit 1 }
Write-Host "PASS"; exit 0
