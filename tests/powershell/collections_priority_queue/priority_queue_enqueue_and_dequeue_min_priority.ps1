# vybe-test: powershell/collections_priority_queue/priority_queue_enqueue_and_dequeue_min_priority
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$pq.Enqueue("LowPriority", 10)
$pq.Enqueue("HighPriority", 1)
$first = $pq.Dequeue()
if ($first -ne "HighPriority" -or $pq.Count -ne 1) { Write-Host "FAIL: PriorityQueue min-heap dequeue failed"; exit 1 }
Write-Host "PASS"; exit 0
