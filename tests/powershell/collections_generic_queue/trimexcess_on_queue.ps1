# vybe-test: powershell/collections_generic_queue/trimexcess_on_queue
$q = [System.Collections.Generic.Queue[int]]::new(100)
$q.Enqueue(1); $q.Enqueue(2)
$q.TrimExcess()
if ($q.Count -ne 2) {
    Write-Host "FAIL: TrimExcess failed"
    exit 1
}
Write-Host "PASS"
exit 0
