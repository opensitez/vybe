# vybe-test: powershell/type_biginteger_arithmetic/multiplication_factorial_overflow_int64
$prod = [bigint]1
for ($i = 1; $i -le 25; $i++) {
    $prod = $prod * [bigint]$i
}
$expected = [bigint]::Parse("15511210043330985984000000")
if ($prod -ne $expected) {
    Write-Host "FAIL: 25! expected $expected, got $prod"
    exit 1
}
Write-Host "PASS"
exit 0
