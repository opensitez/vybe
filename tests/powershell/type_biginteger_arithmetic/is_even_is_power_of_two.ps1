# vybe-test: powershell/type_biginteger_arithmetic/is_even_is_power_of_two
$even = [bigint]1024
$odd = [bigint]1025
if (-not $even.IsEven -or $odd.IsEven) {
    Write-Host "FAIL: IsEven check failed"
    exit 1
}
if (-not $even.IsPowerOfTwo -or $odd.IsPowerOfTwo) {
    Write-Host "FAIL: IsPowerOfTwo check failed"
    exit 1
}
Write-Host "PASS"
exit 0
