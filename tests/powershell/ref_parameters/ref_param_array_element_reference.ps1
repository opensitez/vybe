# vybe-test: powershell/ref_parameters/ref_param_array_element_reference
function Mutate-Element([ref]$elem) {
    $elem.Value = $elem.Value * 10
}
$arr = @(10, 20, 30)
Mutate-Element ([ref]$arr[1])
if ($arr[1] -ne 200) {
    Write-Host "FAIL: array element ref mutation expected 200, got $($arr[1])"
    exit 1
}
Write-Host "PASS"
exit 0
