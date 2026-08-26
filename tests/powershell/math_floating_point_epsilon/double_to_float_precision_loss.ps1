# vybe-test: powershell/math_floating_point_epsilon/double_to_float_precision_loss
$d = 1.23456789012345
$f = [float]$d
$backToDouble = [double]$f
if ($d -eq $backToDouble) {
    Write-Host "FAIL: Cast to float should truncate precision bits"
    exit 1
}
Write-Host "PASS"
exit 0
