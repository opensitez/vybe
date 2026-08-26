# vybe-test: powershell/type_biginteger_arithmetic/bitwise_and_mask
$a = [bigint]::Parse("111111111111111111111111111111")
$b = [bigint]::Parse("000000000000000000000000000015")
$c = $a -band $b
if ($c -ne ([bigint]7)) {
    Write-Host "FAIL: expected 7, got $c"
    exit 1
}
Write-Host "PASS"
exit 0
