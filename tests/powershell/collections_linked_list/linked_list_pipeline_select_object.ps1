# vybe-test: powershell/collections_linked_list/linked_list_pipeline_select_object
$ll = [System.Collections.Generic.LinkedList[string]]::new([string[]]@("hello", "world"))
$res = @($ll | Select-Object @{ N = "Len"; E = { $_.Length } })
if ($res[0].Len -ne 5 -or $res[1].Len -ne 5) { Write-Host "FAIL: Pipeline Select-Object failed"; exit 1 }
Write-Host "PASS"; exit 0
