# vybe-test: powershell/bitwise_operators/not_operator
$val = -bnot 0
if ($val -ne -1) {
    Write-Host "FAIL: bitwise NOT failed"
    exit 1
}
Write-Host "PASS"
exit 0
