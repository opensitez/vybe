# vybe-test: powershell/type_nullable_value_types/get_value_or_default_no_arg
[System.Nullable[int]]$x = $null
if ($x.GetValueOrDefault() -ne 0) {
    Write-Host "FAIL: GetValueOrDefault() should return 0 for default int"
    exit 1
}
Write-Host "PASS"
exit 0
