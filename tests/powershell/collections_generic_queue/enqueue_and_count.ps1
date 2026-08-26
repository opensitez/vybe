# vybe-test: powershell/collections_generic_queue/enqueue_and_count
$q = [System.Collections.Generic.Queue[string]]::new()
$q.Enqueue("first")
$q.Enqueue("second")
if ($q.Count -ne 2) {
    Write-Host "FAIL: Queue Enqueue Count mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
