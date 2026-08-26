# vybe-test: powershell/collections_generic_queue/invalidoperationexception_on_empty_dequeue
$q = [System.Collections.Generic.Queue[int]]::new()
$caught = $false
try {
    $x = $q.Dequeue()
} catch [System.InvalidOperationException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Dequeue on empty queue must throw InvalidOperationException"
    exit 1
}
Write-Host "PASS"
exit 0
