# vybe-test: powershell/token_and_grammar_system/numeric_float_basic
$pi = 3.14
$sum = $pi + 2.0
if ([Math]::Abs($sum - 5.14) -gt 0.000001) {
    Write-Host "FAIL: expected 5.14, got $sum"
    exit 1
}

Write-Host 'PASS'
exit 0
