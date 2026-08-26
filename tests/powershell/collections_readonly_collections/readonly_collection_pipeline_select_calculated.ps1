# vybe-test: powershell/collections_readonly_collections/readonly_collection_pipeline_select_calculated
$list = [System.Collections.Generic.List[string]]::new([string[]]@("hello", "vybe"))
$roc = [System.Collections.ObjectModel.ReadOnlyCollection[string]]::new($list)
$res = @($roc | Select-Object @{ N = "Upper"; E = { $_.ToUpper() } })
if ($res[0].Upper -ne "HELLO" -or $res[1].Upper -ne "VYBE") { Write-Host "FAIL: Pipeline calculated property failed"; exit 1 }
Write-Host "PASS"; exit 0
