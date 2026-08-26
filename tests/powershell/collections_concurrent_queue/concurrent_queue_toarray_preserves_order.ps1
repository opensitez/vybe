# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_toarray_preserves_order
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new()
$cq.Enqueue(10); $cq.Enqueue(20); $cq.Enqueue(30)
$arr = $cq.ToArray()
if ($arr.Length -ne 3 -or $arr[0] -ne 10 -or $arr[2] -ne 30) { Write-Host "FAIL: ToArray failed"; exit 1 }
Write-Host "PASS"; exit 0
