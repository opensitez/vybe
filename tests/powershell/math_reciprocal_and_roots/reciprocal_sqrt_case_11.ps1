# vybe-test: powershell/math_reciprocal_and_roots/reciprocal_sqrt_case_11
$sqrt = [math]::Sqrt(16.0)
$recip = 1.0 / $sqrt
if ($sqrt -ne 4.0 -or $recip -ne 0.25) { Write-Host "FAIL: Sqrt reciprocal failed"; exit 1 }
Write-Host "PASS"; exit 0
