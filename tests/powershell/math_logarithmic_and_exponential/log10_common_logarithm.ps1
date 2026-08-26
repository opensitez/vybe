# vybe-test: powershell/math_logarithmic_and_exponential/log10_common_logarithm
$log100 = [math]::Log10(100.0)
$log1000 = [math]::Log10(1000.0)
if ($log100 -ne 2.0 -or $log1000 -ne 3.0) {
    Write-Host "FAIL: Log10 failed"
    exit 1
}
Write-Host "PASS"
exit 0
