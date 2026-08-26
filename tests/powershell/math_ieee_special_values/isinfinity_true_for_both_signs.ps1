# vybe-test: powershell/math_ieee_special_values/isinfinity_true_for_both_signs
$p = [double]::PositiveInfinity
$n = [double]::NegativeInfinity
if (-not [double]::IsInfinity($p) -or -not [double]::IsInfinity($n)) {
    Write-Host "FAIL: IsInfinity failed"
    exit 1
}
Write-Host "PASS"
exit 0
