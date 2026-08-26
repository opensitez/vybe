# vybe-test: powershell/type_nullable_value_types/equality_two_nullables_same_value
[System.Nullable[int]]$a = 50
[System.Nullable[int]]$b = 50
if ($a -ne $b) {
    Write-Host "FAIL: two nullables with same value must be equal"
    exit 1
}
Write-Host "PASS"
exit 0
