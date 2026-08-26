# vybe-test: powershell/collections_immutable_arrays/immutable_array_measure_object_sum
$arr = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(1, 2, 3, 4))
$m = $arr | Measure-Object -Sum
if ($m.Sum -ne 10) { Write-Host "FAIL: Measure-Object failed"; exit 1 }
Write-Host "PASS"; exit 0
