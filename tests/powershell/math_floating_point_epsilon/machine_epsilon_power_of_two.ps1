# vybe-test: powershell/math_floating_point_epsilon/machine_epsilon_power_of_two
# 2^-52 for 64-bit float
$machEps = [math]::Pow(2.0, -52.0)
if ($machEps -lt 2.22e-16 -or $machEps -gt 2.23e-16) {
    Write-Host "FAIL: Machine epsilon 2^-52 failed, got $machEps"
    exit 1
}
Write-Host "PASS"
exit 0
