# vybe-test: powershell/collections_generic_queue/invalidoperationexception_on_empty_peek
$q = [System.Collections.Generic.Queue[int]]::new()
$caught = $false
try {
    $x = $q.Peek()
} catch [System.InvalidOperationException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Peek on empty queue must throw InvalidOperationException"
    exit 1
}
Write-Host "PASS"
exit 0
