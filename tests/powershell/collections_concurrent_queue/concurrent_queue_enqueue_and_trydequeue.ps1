# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_enqueue_and_trydequeue
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new()
$cq.Enqueue(100)
[int]$outVal = 0
$ok = $cq.TryDequeue([ref]$outVal)
if (-not $ok -or $outVal -ne 100 -or $cq.Count -ne 0) { Write-Host "FAIL: Enqueue/TryDequeue failed"; exit 1 }
Write-Host "PASS"; exit 0
