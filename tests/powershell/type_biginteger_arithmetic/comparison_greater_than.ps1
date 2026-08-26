# vybe-test: powershell/type_biginteger_arithmetic/comparison_greater_than
$a = [bigint]::Parse("100000000000000000000000000001")
$b = [bigint]::Parse("100000000000000000000000000000")
if (-not ($a -gt $b)) {
    Write-Host "FAIL: $a should be greater than $b"
    exit 1
}
Write-Host "PASS"
exit 0
