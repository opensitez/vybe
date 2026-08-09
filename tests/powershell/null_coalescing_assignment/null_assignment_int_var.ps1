# vybe-test: powershell/null_coalescing_assignment/null_assignment_int_var
$intVar = $null
$intVar ??= 500
if ($intVar -ne 500) {
    Write-Host "FAIL: int variable ??= expected 500, got $intVar"
    exit 1
}
Write-Host "PASS"
exit 0
