# vybe-test: powershell/type_nullable_value_types/equality_nullable_with_literal_null
[System.Nullable[int]]$a = $null
if ($a -ne $null) {
    Write-Host "FAIL: null nullable must equal $null"
    exit 1
}
Write-Host "PASS"
exit 0
