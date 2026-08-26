# vybe-test: powershell/math_floating_point_epsilon/nearly_equal_helper_function
function Assert-NearlyEqual([double]$actual, [double]$expected, [double]$tolerance = 1e-7) {
    return ([math]::Abs($actual - $expected) -le $tolerance)
}
$ok = Assert-NearlyEqual (0.1 + 0.2) 0.3
if (-not $ok) {
    Write-Host "FAIL: NearlyEqual helper failed"
    exit 1
}
Write-Host "PASS"
exit 0
