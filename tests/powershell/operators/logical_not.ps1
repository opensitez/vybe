# vybe-test: powershell/operators/logical_not
$result = (-not $true)
if ($result -ne $false) {
    Write-Host "FAIL: expected False, got $result"
    exit 1
}
$result2 = (-not $false)
if ($result2 -ne $true) {
    Write-Host "FAIL: expected True, got $result2"
    exit 1
}
Write-Host "PASS"
exit 0
