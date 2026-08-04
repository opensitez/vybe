# vybe-test: powershell/comparison/not_equal_operator
$result = (5 -ne 3)
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
