# vybe-test: powershell/ref_parameters/ref_param_return_and_ref_combination
$x = 10
$x += 5
$x *= 2
if ($x -eq 30) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
