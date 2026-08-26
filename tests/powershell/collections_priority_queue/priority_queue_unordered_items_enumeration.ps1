# vybe-test: powershell/collections_priority_queue/priority_queue_unordered_items_enumeration
$pq = [System.Collections.Generic.PriorityQueue[string, int]]::new()
$pq.Enqueue("A", 1); $pq.Enqueue("B", 2)
$items = @($pq.UnorderedItems)
if ($items.Length -ne 2) { Write-Host "FAIL: UnorderedItems enumeration failed"; exit 1 }
Write-Host "PASS"; exit 0
