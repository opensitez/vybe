# vybe-test: powershell/language_ternary_conditional_operator/ternary_nested_conditions
$score = 85
$grade = ($score -ge 90) ? "A" : (($score -ge 80) ? "B" : "C")
if ($grade -ne "B") {
    Write-Host "FAIL: Nested ternary conditions failed, got '$grade'"
    exit 1
}
Write-Host "PASS"
exit 0
