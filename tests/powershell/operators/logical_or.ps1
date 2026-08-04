# vybe-test: powershell/operators/logical_or
$result = ($true -or $false)
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
$result2 = ($false -or $false)
if ($result2 -ne $false) {
    Write-Host "FAIL: expected False, got $result2"
    exit 1
}
Write-Host "PASS"
exit 0
