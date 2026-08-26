# vybe-test: powershell/language_ternary_conditional_operator/ternary_in_class_method
class TernaryClass {
    [string]CheckNum([int]$n) {
        return ($n -ge 0) ? "PositiveOrZero" : "Negative"
    }
}
$tc = [TernaryClass]::new()
if ($tc.CheckNum(5) -ne "PositiveOrZero" -or $tc.CheckNum(-3) -ne "Negative") {
    Write-Host "FAIL: Ternary in class method failed"
    exit 1
}
Write-Host "PASS"
exit 0
