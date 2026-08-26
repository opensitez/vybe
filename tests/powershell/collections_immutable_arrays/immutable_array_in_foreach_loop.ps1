# vybe-test: powershell/collections_immutable_arrays/immutable_array_in_foreach_loop
$arr = [System.Collections.Immutable.ImmutableArray]::Create([int[]]@(10, 20, 30))
$sum = 0
foreach ($item in $arr) { $sum += $item }
if ($sum -ne 60) { Write-Host "FAIL: foreach failed"; exit 1 }
Write-Host "PASS"; exit 0
