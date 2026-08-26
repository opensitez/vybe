# vybe-test: powershell/collections_priority_queue/priority_queue_while_loop_drain
$pq = [System.Collections.Generic.PriorityQueue[int, int]]::new()
$pq.Enqueue(3, 3); $pq.Enqueue(1, 1); $pq.Enqueue(2, 2)
$list = [System.Collections.Generic.List[int]]::new()
while ($pq.Count -gt 0) { $list.Add($pq.Dequeue()) }
if ($list[0] -ne 1 -or $list[1] -ne 2 -or $list[2] -ne 3) { Write-Host "FAIL: Drain while loop failed"; exit 1 }
Write-Host "PASS"; exit 0
