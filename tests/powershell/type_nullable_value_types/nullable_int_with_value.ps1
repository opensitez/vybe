# vybe-test: powershell/type_nullable_value_types/nullable_int_with_value
$n = [System.Activator]::CreateInstance([type]"System.Nullable[int]", @(42))
if ($n -ne 42) {
    Write-Host "FAIL: Nullable int with value failed"
    exit 1
}
Write-Host "PASS"
exit 0
