# vybe-test: powershell/comparison/notin_operator
$result = 10 -notin @(1, 2, 3, 4, 5)
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
