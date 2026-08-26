# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_count_consistency
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new()
for ($i = 0; $i -lt 50; $i++) { $cq.Enqueue($i) }
if ($cq.Count -ne 50) { Write-Host "FAIL: Count consistency failed"; exit 1 }
Write-Host "PASS"; exit 0
