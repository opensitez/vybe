# vybe-test: powershell/collections_generic_queue/interleaved_enqueue_dequeue
$q = [System.Collections.Generic.Queue[int]]::new()
$q.Enqueue(1)
$q.Enqueue(2)
$d1 = $q.Dequeue() # 1
$q.Enqueue(3)
$d2 = $q.Dequeue() # 2
$d3 = $q.Dequeue() # 3
if ($d1 -ne 1 -or $d2 -ne 2 -or $d3 -ne 3 -or $q.Count -ne 0) {
    Write-Host "FAIL: Interleaved enqueue/dequeue failed"
    exit 1
}
Write-Host "PASS"
exit 0
