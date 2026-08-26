# vybe-test: powershell/collections_generic_queue/trydequeue_success_and_empty
$q = [System.Collections.Generic.Queue[int]]::new()
$q.Enqueue(100)
$hasItem = ($q.Count -gt 0)
$item = if ($hasItem) { $q.Dequeue() } else { 0 }
if (-not $hasItem -or $item -ne 100 -or $q.Count -ne 0) {
    Write-Host "FAIL: Queue conditional dequeue failed"
    exit 1
}
Write-Host "PASS"
exit 0
