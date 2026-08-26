# vybe-test: powershell/language_ternary_conditional_operator/ternary_returning_numeric_calculations
$x = 10
$y = ($x -gt 5) ? ($x * 2) : ($x / 2)
if ($y -ne 20) {
    Write-Host "FAIL: Ternary returning numeric calculation failed, got $y"
    exit 1
}
Write-Host "PASS"
exit 0
