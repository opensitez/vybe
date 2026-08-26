# vybe-test: powershell/math_logarithmic_and_exponential/math_e_constant_precision
$e = [math]::E
if ($e -lt 2.71828182845904 -or $e -gt 2.71828182845905) {
    Write-Host "FAIL: Math::E precision failed"
    exit 1
}
Write-Host "PASS"
exit 0
