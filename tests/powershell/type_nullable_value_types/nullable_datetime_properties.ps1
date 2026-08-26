# vybe-test: powershell/type_nullable_value_types/nullable_datetime_properties
$dt = [datetime]::Parse("2026-08-26")
$t = [type]"System.Nullable[datetime]"
$inst = [System.Activator]::CreateInstance($t, @($dt))
$val = $t.GetProperty("Value").GetValue($inst)
if ($val.Year -ne 2026 -or $val.Month -ne 8 -or $val.Day -ne 26) {
    Write-Host "FAIL: Nullable datetime property extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
