# vybe-test: powershell/ref_parameters/ref_param_in_custom_function
function Set-Double([ref]$x) {
    $x.Value = $x.Value * 2
}
$n = 25
Set-Double -x ([ref]$n)
if ($n -ne 50) {
    Write-Host "FAIL: named parameter ref mutation expected 50, got $n"
    exit 1
}
Write-Host "PASS"
exit 0
