# vybe-test: powershell/arithmetic_operators/exponentiation
$res = [math]::Pow(2, 8)
if ($res -ne 256) {
    Write-Host "FAIL: Exponentiation failed"
    exit 1
}
Write-Host "PASS"
exit 0
