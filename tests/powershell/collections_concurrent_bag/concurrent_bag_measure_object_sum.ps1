# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_measure_object_sum
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new([int[]]@(1, 2, 3, 4))
$m = $bag | Measure-Object -Sum
if ($m.Sum -ne 10) { Write-Host "FAIL: Measure-Object on ConcurrentBag failed"; exit 1 }
Write-Host "PASS"; exit 0
