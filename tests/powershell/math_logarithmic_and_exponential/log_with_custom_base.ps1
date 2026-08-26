# vybe-test: powershell/math_logarithmic_and_exponential/log_with_custom_base
$logBase3 = [math]::Log(81.0, 3.0)
if ([math]::Abs($logBase3 - 4.0) -gt 1e-12) {
    Write-Host "FAIL: Log with base 3 expected 4, got $logBase3"
    exit 1
}
Write-Host "PASS"
exit 0
