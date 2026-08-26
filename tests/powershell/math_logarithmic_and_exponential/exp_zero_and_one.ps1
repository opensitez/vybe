# vybe-test: powershell/math_logarithmic_and_exponential/exp_zero_and_one
$e0 = [math]::Exp(0.0)
$e1 = [math]::Exp(1.0)
if ($e0 -ne 1.0 -or [math]::Abs($e1 - [math]::E) -gt 1e-12) {
    Write-Host "FAIL: Exp(0) / Exp(1) failed"
    exit 1
}
Write-Host "PASS"
exit 0
