# vybe-test: powershell/math_trigonometric_functions/trig_function_symmetry_odd_even
$x = 0.75
$sinPos = [math]::Sin($x)
$sinNeg = [math]::Sin(-$x)
$cosPos = [math]::Cos($x)
$cosNeg = [math]::Cos(-$x)
if ([math]::Abs($sinPos - (-$sinNeg)) -gt 1e-12 -or [math]::Abs($cosPos - $cosNeg) -gt 1e-12) {
    Write-Host "FAIL: Odd/Even symmetry of Sin/Cos failed"
    exit 1
}
Write-Host "PASS"
exit 0
