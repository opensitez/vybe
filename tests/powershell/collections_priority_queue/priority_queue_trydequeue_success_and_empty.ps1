# vybe-test: powershell/collections_priority_queue/priority_queue_trydequeue_success_and_empty
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$pq.Enqueue("Task1", 5)
$hasItem = ($pq.Count -gt 0)
$item = if ($hasItem) { $pq.Dequeue() } else { "" }
if (-not $hasItem -or $item -ne "Task1" -or $pq.Count -ne 0) { Write-Host "FAIL: Priority queue dequeue failed"; exit 1 }
Write-Host "PASS"; exit 0
