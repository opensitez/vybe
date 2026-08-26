# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_is_empty_property
$cq = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
if (-not $cq.IsEmpty) { Write-Host "FAIL: Initial IsEmpty should be true"; exit 1 }
$cq.Enqueue("item")
if ($cq.IsEmpty) { Write-Host "FAIL: IsEmpty after Enqueue should be false"; exit 1 }
Write-Host "PASS"; exit 0
