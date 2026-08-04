# vybe-test: powershell/comparison/less_than_operator
$result = (3 -lt 7)
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
