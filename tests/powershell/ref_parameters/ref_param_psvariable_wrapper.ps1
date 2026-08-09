# vybe-test: powershell/ref_parameters/ref_param_psvariable_wrapper
$v = 500
$ref = [ref]$v
if (-not ($ref.Value -is [int])) {
    Write-Host "FAIL: ref Value is not [int]"
    exit 1
}
if ($ref.GetType().Name -ne "PSReference") {
    Write-Host "FAIL: ref object type expected PSReference, got $($ref.GetType().Name)"
    exit 1
}
Write-Host "PASS"
exit 0
