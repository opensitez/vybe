# vybe-test: powershell/numeric_overflow_and_precision/basic
$max = [double]::MaxValue
$overflow = $max * 2

if (-not [double]::IsInfinity($overflow)) {
    Write-Host "FAIL: expected overflow to infinity, got $overflow"
    exit 1
}

$sum = 0.1 + 0.2
if ([Math]::Abs($sum - 0.3) -gt 0.0000001) {
    Write-Host "FAIL: floating precision not as expected, got $sum"
    exit 1
}

Write-Host 'PASS'
exit 0
