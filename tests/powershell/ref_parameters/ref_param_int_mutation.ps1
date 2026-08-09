# vybe-test: powershell/ref_parameters/ref_param_int_mutation
function Increment-Ref([ref]$num) {
    $num.Value++
}
$val = 10
Increment-Ref ([ref]$val)
if ($val -ne 11) {
    Write-Host "FAIL: [ref] int increment expected 11, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
