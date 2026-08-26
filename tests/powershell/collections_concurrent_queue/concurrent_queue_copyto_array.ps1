# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_copyto_array
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new()
$cq.Enqueue(1); $cq.Enqueue(2)
[int[]]$arr = [int[]]::new(2)
$cq.CopyTo($arr, 0)
if ($arr[0] -ne 1 -or $arr[1] -ne 2) { Write-Host "FAIL: CopyTo failed"; exit 1 }
Write-Host "PASS"; exit 0
