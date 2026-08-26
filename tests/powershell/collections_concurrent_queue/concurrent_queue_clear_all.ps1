# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_clear_all
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new()
$cq.Enqueue(1); $cq.Enqueue(2)
$cq.Clear()
if ($cq.Count -ne 0 -or -not $cq.IsEmpty) { Write-Host "FAIL: Clear failed"; exit 1 }
Write-Host "PASS"; exit 0
