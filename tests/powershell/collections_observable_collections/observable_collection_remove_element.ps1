# vybe-test: powershell/collections_observable_collections/observable_collection_remove_element
$oc = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
$oc.Add("a"); $oc.Add("b")
$rem = $oc.Remove("a")
if (-not $rem -or $oc.Count -ne 1 -or $oc[0] -ne "b") { Write-Host "FAIL: ObservableCollection Remove failed"; exit 1 }
Write-Host "PASS"; exit 0
