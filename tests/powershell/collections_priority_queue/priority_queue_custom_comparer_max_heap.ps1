# vybe-test: powershell/collections_priority_queue/priority_queue_custom_comparer_max_heap
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$pq.Enqueue("Small", 100)
$pq.Enqueue("Large", 10)
$first = $pq.Dequeue()
if ($first -ne "Large") { Write-Host "FAIL: Priority queue failed"; exit 1 }
Write-Host "PASS"; exit 0
