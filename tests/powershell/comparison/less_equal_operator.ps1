# vybe-test: powershell/comparison/less_equal_operator
$result = (5 -le 5)
if ($result -ne $true) {
    Write-Host "FAIL: expected True for 5 <= 5"
    exit 1
}
$result2 = (4 -le 5)
if ($result2 -ne $true) {
    Write-Host "FAIL: expected True for 4 <= 5"
    exit 1
}
Write-Host "PASS"
exit 0
