# vybe-test: powershell/math_logarithmic_and_exponential/log2_binary_logarithm
$l8 = [math]::Log2(8.0)
$l1024 = [math]::Log2(1024.0)
if ($l8 -ne 3.0 -or $l1024 -ne 10.0) {
    Write-Host "FAIL: Log2 failed"
    exit 1
}
Write-Host "PASS"
exit 0
