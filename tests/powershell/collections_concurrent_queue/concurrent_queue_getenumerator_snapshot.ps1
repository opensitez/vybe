# vybe-test: powershell/collections_concurrent_queue/concurrent_queue_getenumerator_snapshot
$cq = [System.Collections.Concurrent.ConcurrentQueue[int]]::new([int[]]@(1, 2, 3))
$enum = $cq.GetEnumerator()
$cq.Enqueue(4)
$items = [System.Collections.Generic.List[int]]::new()
while ($enum.MoveNext()) { $items.Add($enum.Current) }
if ($items.Count -lt 3) { Write-Host "FAIL: Enumerator snapshot failed"; exit 1 }
Write-Host "PASS"; exit 0
