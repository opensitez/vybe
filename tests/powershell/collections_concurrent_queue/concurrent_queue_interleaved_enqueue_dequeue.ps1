# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_interleaved_enqueue_dequeue
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new()
$cq.Enqueue(1)
[int]$outVal = 0
$null = $cq.TryDequeue([ref]$outVal)
$cq.Enqueue(2)
$null = $cq.TryDequeue([ref]$outVal)
if ($outVal -ne 2 -or $cq.Count -ne 0) { Write-Host "FAIL: Interleaved enqueue/dequeue failed"; exit 1 }
Write-Host "PASS"; exit 0
