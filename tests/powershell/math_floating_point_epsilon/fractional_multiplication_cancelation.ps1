# vybe-test: powershell/math_floating_point_epsilon/fractional_multiplication_cancelation
$x = 1e16
$res = ($x + 1.0) - $x
# In double precision, 1e16 + 1.0 loses the 1.0 bit
if ($res -ne 0.0 -and $res -ne 1.0) {
    Write-Host "FAIL: Catastrophic cancellation behavior check"
    exit 1
}
Write-Host "PASS"
exit 0
