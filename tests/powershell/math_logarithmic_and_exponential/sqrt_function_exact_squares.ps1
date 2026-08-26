# vybe-test: powershell/math_logarithmic_and_exponential/sqrt_function_exact_squares
$s1 = [math]::Sqrt(9.0)
$s2 = [math]::Sqrt(144.0)
if ($s1 -ne 3.0 -or $s2 -ne 12.0) {
    Write-Host "FAIL: Sqrt exact squares failed"
    exit 1
}
Write-Host "PASS"
exit 0
