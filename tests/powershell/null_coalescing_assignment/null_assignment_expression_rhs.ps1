# vybe-test: powershell/null_coalescing_assignment/null_assignment_expression_rhs
$data = $null
$data ??= (10 + 20)
if ($data -ne 30) {
    Write-Host "FAIL: expression RHS ??= expected 30, got $data"
    exit 1
}
Write-Host "PASS"
exit 0
