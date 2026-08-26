# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_fifo_ordering
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new()
$cq.Enqueue(1); $cq.Enqueue(2); $cq.Enqueue(3)
[int]$v1 = 0; [int]$v2 = 0; [int]$v3 = 0
$null = $cq.TryDequeue([ref]$v1)
$null = $cq.TryDequeue([ref]$v2)
$null = $cq.TryDequeue([ref]$v3)
if ($v1 -ne 1 -or $v2 -ne 2 -or $v3 -ne 3) { Write-Host "FAIL: FIFO ordering failed"; exit 1 }
Write-Host "PASS"; exit 0
