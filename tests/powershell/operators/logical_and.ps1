# vybe-test: powershell/operators/logical_and
$result = ($true -and $true)
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
$result2 = ($true -and $false)
if ($result2 -ne $false) {
    Write-Host "FAIL: expected False, got $result2"
    exit 1
}
Write-Host "PASS"
exit 0
