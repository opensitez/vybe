# vybe-test: powershell/type_nullable_value_types/nullable_double_precision
$t = [type]"System.Nullable[double]"
$inst = [System.Activator]::CreateInstance($t, @(123.456789))
$val = $t.GetProperty("Value").GetValue($inst)
if ($val -ne 123.456789) {
    Write-Host "FAIL: Nullable double precision mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
