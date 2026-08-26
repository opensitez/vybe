# vybe-test: powershell/math_logarithmic_and_exponential/log_multiplication_property
# log(a * b) = log(a) + log(b)
$a = 5.0
$b = 7.0
$lhs = [math]::Log($a * $b)
$rhs = [math]::Log($a) + [math]::Log($b)
if ([math]::Abs($lhs - $rhs) -gt 1e-12) {
    Write-Host "FAIL: Log multiplication rule failed"
    exit 1
}
Write-Host "PASS"
exit 0
