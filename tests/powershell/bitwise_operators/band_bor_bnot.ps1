# vybe-test: powershell/bitwise_operators/band_bor_bnot
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
