# vybe-test: powershell/type_nullable_value_types/equality_nullable_with_raw_value
[System.Nullable[int]]$a = 25
if ($a -ne 25) {
    Write-Host "FAIL: nullable(25) must equal raw 25"
    exit 1
}
Write-Host "PASS"
exit 0
