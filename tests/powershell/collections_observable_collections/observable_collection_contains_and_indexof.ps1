# vybe-test: powershell/collections_observable_collections/observable_collection_contains_and_indexof
$oc = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
$oc.Add("cat"); $oc.Add("dog")
if (-not $oc.Contains("dog") -or $oc.IndexOf("dog") -ne 1) { Write-Host "FAIL: Contains/IndexOf failed"; exit 1 }
Write-Host "PASS"; exit 0
