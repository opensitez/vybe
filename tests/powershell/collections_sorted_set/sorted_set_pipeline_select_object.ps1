# vybe-test: powershell/collections_sorted_set/sorted_set_pipeline_select_object
$ss = [System.Collections.Generic.SortedSet[string]]::new([string[]]@("b", "a"))
$res = @($ss | Select-Object @{ N = "Upper"; E = { $_.ToUpper() } })
if ($res[0].Upper -ne "A" -or $res[1].Upper -ne "B") { Write-Host "FAIL: Pipeline Select-Object failed"; exit 1 }
Write-Host "PASS"; exit 0
