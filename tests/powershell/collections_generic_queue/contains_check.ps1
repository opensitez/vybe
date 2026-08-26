# vybe-test: powershell/collections_generic_queue/contains_check
$q = [System.Collections.Generic.Queue[int]]::new()
$q.Enqueue(1); $q.Enqueue(2)
if (-not $q.Contains(2) -or $q.Contains(5)) {
    Write-Host "FAIL: Queue Contains failed"
    exit 1
}
Write-Host "PASS"
exit 0
