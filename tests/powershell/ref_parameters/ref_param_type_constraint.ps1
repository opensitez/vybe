# vybe-test: powershell/ref_parameters/ref_param_type_constraint
[int]$typedVar = 50
$r = [ref]$typedVar
$r.Value = "100"
if ($typedVar -ne 100 -or -not ($typedVar -is [int])) {
    Write-Host "FAIL: type-constrained ref variable expected int 100"
    exit 1
}
Write-Host "PASS"
exit 0
