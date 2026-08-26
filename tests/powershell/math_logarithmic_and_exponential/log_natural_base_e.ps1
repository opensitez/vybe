# vybe-test: powershell/math_logarithmic_and_exponential/log_natural_base_e
$logE = [math]::Log([math]::E)
$log1 = [math]::Log(1.0)
if ([math]::Abs($logE - 1.0) -gt 1e-12 -or $log1 -ne 0.0) {
    Write-Host "FAIL: Natural Log failed"
    exit 1
}
Write-Host "PASS"
exit 0
