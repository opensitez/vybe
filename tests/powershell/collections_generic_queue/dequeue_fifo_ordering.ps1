# vybe-test: powershell/collections_generic_queue/dequeue_fifo_ordering
$q = [System.Collections.Generic.Queue[int]]::new()
$q.Enqueue(10)
$q.Enqueue(20)
$q.Enqueue(30)
$item1 = $q.Dequeue()
$item2 = $q.Dequeue()
if ($item1 -ne 10 -or $item2 -ne 20 -or $q.Count -ne 1) {
    Write-Host "FAIL: Queue Dequeue FIFO ordering failed"
    exit 1
}
Write-Host "PASS"
exit 0
