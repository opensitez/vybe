# vybe-test: powershell/collections_observable_collections/observable_collection_move_element
$oc = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
$oc.Add("first"); $oc.Add("second"); $oc.Add("third")
$oc.Move(0, 2)
if ($oc[0] -ne "second" -or $oc[2] -ne "first") { Write-Host "FAIL: Move failed"; exit 1 }
Write-Host "PASS"; exit 0
