# vybe-test: powershell/math_trigonometric_functions/math_pi_constant_precision
$pi = [math]::PI
if ($pi -lt 3.14159265358979 -or $pi -gt 3.14159265358980) {
    Write-Host "FAIL: Math::PI precision failed"
    exit 1
}
Write-Host "PASS"
exit 0
