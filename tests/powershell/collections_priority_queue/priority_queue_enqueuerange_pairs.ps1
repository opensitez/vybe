# vybe-test: powershell/collections_priority_queue/priority_queue_enqueuerange_pairs
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$p1 = [System.Tuple[string, int]]::new("Third", 30)
$p2 = [System.Tuple[string, int]]::new("First", 10)
$p3 = [System.Tuple[string, int]]::new("Second", 20)
$pq.EnqueueRange(@($p1, $p2, $p3))
if ($pq.Dequeue() -ne "First" -or $pq.Dequeue() -ne "Second" -or $pq.Dequeue() -ne "Third") {
    Write-Host "FAIL: EnqueueRange failed"; exit 1
}
Write-Host "PASS"; exit 0
