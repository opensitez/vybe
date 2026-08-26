# vybe-test: powershell/type_nullable_value_types/nullable_int_null_assignment
[System.Nullable[int]]$n = $null
if ($n.HasValue) {
    Write-Host "FAIL: Nullable int assigned null should have HasValue = false"
    exit 1
}
Write-Host "PASS"
exit 0
