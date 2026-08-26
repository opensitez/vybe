# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_in_foreach_loop
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new([int[]]@(10, 20, 30))
$sum = 0
foreach ($item in $cq) { $sum += $item }
if ($sum -ne 60) { Write-Host "FAIL: foreach on ConcurrentQueue failed"; exit 1 }
Write-Host "PASS"; exit 0
