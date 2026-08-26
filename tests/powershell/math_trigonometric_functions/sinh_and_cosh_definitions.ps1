# vybe-test: powershell/math_trigonometric_functions/sinh_and_cosh_definitions
$x = 1.0
$sinh = [math]::Sinh($x)
$cosh = [math]::Cosh($x)
$diff = ($cosh * $cosh) - ($sinh * $sinh) # cosh^2 - sinh^2 = 1
if ([math]::Abs($diff - 1.0) -gt 1e-12) {
    Write-Host "FAIL: Hyperbolic identity cosh^2 - sinh^2 expected 1.0, got $diff"
    exit 1
}
Write-Host "PASS"
exit 0
