# vybe-test: powershell/math_trigonometric_functions/trigonometric_angle_addition_formula
$a = 0.3
$b = 0.4
$sinAplusB = [math]::Sin($a + $b)
$formula = ([math]::Sin($a) * [math]::Cos($b)) + ([math]::Cos($a) * [math]::Sin($b))
if ([math]::Abs($sinAplusB - $formula) -gt 1e-12) {
    Write-Host "FAIL: Angle addition formula failed"
    exit 1
}
Write-Host "PASS"
exit 0
