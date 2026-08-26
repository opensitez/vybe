# vybe-test: powershell/math_trigonometric_functions/pythagorean_identity_sin2_cos2
$theta = 1.2345
$sin = [math]::Sin($theta)
$cos = [math]::Cos($theta)
$identity = ($sin * $sin) + ($cos * $cos)
if ([math]::Abs($identity - 1.0) -gt 1e-12) {
    Write-Host "FAIL: sin^2 + cos^2 expected 1.0, got $identity"
    exit 1
}
Write-Host "PASS"
exit 0
