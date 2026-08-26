# vybe-test: powershell/collections_observable_collections/observable_collection_with_complex_objects
$oc = [System.Collections.ObjectModel.ObservableCollection[pscustomobject]]::new()
$oc.Add([pscustomobject]@{ Id = 1; Label = "Item1" })
if ($oc[0].Label -ne "Item1") { Write-Host "FAIL: Complex object failed"; exit 1 }
Write-Host "PASS"; exit 0
