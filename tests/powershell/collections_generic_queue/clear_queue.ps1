# vybe-test: powershell/collections_generic_queue/clear_queue
$q = [System.Collections.Generic.Queue[int]]::new()
$q.Enqueue(1); $q.Enqueue(2)
$q.Clear()
if ($q.Count -ne 0) {
    Write-Host "FAIL: Queue Clear failed"
    exit 1
}
Write-Host "PASS"
exit 0
