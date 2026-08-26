# vybe-test: powershell/collections_linked_list/linked_list_pipeline_filter
$ll = [System.Collections.Generic.LinkedList[int]]::new([int[]]@(1..5))
$gt3 = @($ll | Where-Object { $_ -gt 3 })
if ($gt3.Length -ne 2 -or $gt3[0] -ne 4 -or $gt3[1] -ne 5) { Write-Host "FAIL: Pipeline filter failed"; exit 1 }
Write-Host "PASS"; exit 0
