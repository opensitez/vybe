# vybe-test: powershell/collections_concurrent_bag/concurrent_bag_in_foreach_loop
$bag = [System.Collections.Concurrent.ConcurrentBag[int]]::new([int[]]@(10, 20, 30))
$sum = 0
foreach ($item in $bag) { $sum += $item }
if ($sum -ne 60) { Write-Host "FAIL: foreach on ConcurrentBag failed"; exit 1 }
Write-Host "PASS"; exit 0
