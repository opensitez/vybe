# vybe-test: powershell/comparison/greater_than_operator
$result = (10 -gt 5)
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
