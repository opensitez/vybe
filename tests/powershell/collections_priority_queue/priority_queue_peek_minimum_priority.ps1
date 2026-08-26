# vybe-test: powershell/collections_priority_queue/priority_queue_peek_minimum_priority
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$pq.Enqueue("B", 2)
$pq.Enqueue("A", 1)
$top = $pq.Peek()
if ($top -ne "A" -or $pq.Count -ne 2) { Write-Host "FAIL: PriorityQueue Peek failed"; exit 1 }
Write-Host "PASS"; exit 0
