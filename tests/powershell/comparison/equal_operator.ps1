# vybe-test: powershell/comparison/equal_operator
$result = (5 -eq 5)
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
