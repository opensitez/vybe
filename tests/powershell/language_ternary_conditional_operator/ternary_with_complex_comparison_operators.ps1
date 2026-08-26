# vybe-test: powershell/language_ternary_conditional_operator/ternary_with_complex_comparison_operators
$val = 50
$res = ($val -gt 20 -and $val -lt 100) ? "InRange" : "OutOfRange"
if ($res -ne "InRange") {
    Write-Host "FAIL: Ternary with complex comparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
