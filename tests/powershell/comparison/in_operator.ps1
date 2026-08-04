# vybe-test: powershell/comparison/in_operator
$result = 3 -in @(1, 2, 3, 4, 5)
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
