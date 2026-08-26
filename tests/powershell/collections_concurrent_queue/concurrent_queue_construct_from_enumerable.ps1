# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_construct_from_enumerable
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new([int[]]@(1, 2, 3, 4))
if ($cq.Count -ne 4) { Write-Host "FAIL: Constructor from enumerable failed"; exit 1 }
Write-Host "PASS"; exit 0
