# vybe-test: powershell/collections_generic_queue/queue_equality_reference
$q1 = [System.Collections.Generic.Queue[int]]::new()
$q2 = $q1
if ($q1 -ne $q2) {
    Write-Host "FAIL: Same queue instance reference equality failed"
    exit 1
}
Write-Host "PASS"
exit 0
