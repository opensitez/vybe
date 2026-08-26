# vybe-test: powershell/math_floating_point_epsilon/accumulated_rounding_error_in_sum_loop
$sum = 0.0
for ($i = 0; $i -lt 10; $i++) { $sum += 0.1 }
# 10 * 0.1 in float != 1.0
$diff = [math]::Abs($sum - 1.0)
if ($diff -gt 1e-14) {
    Write-Host "FAIL: Accumulated rounding delta exceeded tolerance"
    exit 1
}
Write-Host "PASS"
exit 0
