# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_with_null_elements
$cq = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$cq.Enqueue("A"); $cq.Enqueue("B")
[string]$outVal = ""
$null = $cq.TryDequeue([ref]$outVal)
$null = $cq.TryDequeue([ref]$outVal)
if ($outVal -ne "B" -or $cq.Count -ne 0) { Write-Host "FAIL: Queue dequeue failed"; exit 1 }
Write-Host "PASS"; exit 0
