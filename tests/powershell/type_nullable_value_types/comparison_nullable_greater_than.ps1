# vybe-test: powershell/type_nullable_value_types/comparison_nullable_greater_than
[System.Nullable[int]]$a = 20
[System.Nullable[int]]$b = 10
if (-not ($a -gt $b)) {
    Write-Host "FAIL: Nullable comparison failed"
    exit 1
}
Write-Host "PASS"
exit 0
