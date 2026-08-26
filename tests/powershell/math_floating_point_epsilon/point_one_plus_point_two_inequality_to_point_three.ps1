# vybe-test: powershell/math_floating_point_epsilon/point_one_plus_point_two_inequality_to_point_three
$sum = 0.1 + 0.2
$exact = 0.3
# IEEE 754 precision quirk: 0.1 + 0.2 != 0.3 in exact floating point
if ($sum -eq $exact) {
    Write-Host "FAIL: 0.1 + 0.2 should exhibit standard IEEE 754 precision delta"
    exit 1
}
Write-Host "PASS"
exit 0
