# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_trydequeue_empty_returns_false
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new()
[int]$outVal = 0
$ok = $cq.TryDequeue([ref]$outVal)
if ($ok) { Write-Host "FAIL: TryDequeue empty should return false"; exit 1 }
Write-Host "PASS"; exit 0
