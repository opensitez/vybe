# vybe-test: powershell/collections_observable_collections/observable_collection_pipeline_select_calculated
$oc = [System.Collections.ObjectModel.ObservableCollection[string]]::new([string[]]@("a", "b"))
$res = @($oc | Select-Object @{ N = "Upper"; E = { $_.ToUpper() } })
if ($res[0].Upper -ne "A") { Write-Host "FAIL: Pipeline calculated failed"; exit 1 }
Write-Host "PASS"; exit 0
