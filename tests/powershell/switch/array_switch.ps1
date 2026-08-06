# vybe-test: powershell/switch/array_switch
# A switch over a COLLECTION tests every element in turn, so it outputs one
# result per element — not a single value. Here that is notfound, found,
# notfound.
$values = 1, 2, 3
$result = switch ($values) { 2 { 'found' } default { 'notfound' } }
if ($result.Count -ne 3) {
    Write-Host "FAIL: expected one result per element, got $($result.Count)"
    exit 1
}
if ($result[1] -ne 'found') {
    Write-Host "FAIL: element 2 should have matched, got $($result[1])"
    exit 1
}
Write-Host 'PASS'
exit 0
