# vybe-test: powershell/scope/local_scope
$x = 10
function Test-LocalScope {
    $x = 20
    return $x
}
$result = Test-LocalScope
if ($result -ne 20) {
    Write-Host "FAIL: expected 20, got $result"
    exit 1
}
if ($x -ne 10) {
    Write-Host "FAIL: expected outer x = 10, got $x"
    exit 1
}
Write-Host "PASS"
exit 0
