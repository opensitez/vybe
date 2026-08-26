# vybe-test: powershell/type_nullable_value_types/nullable_guid_assignment
$g = [guid]::NewGuid()
$t = [type]"System.Nullable[guid]"
$inst = [System.Activator]::CreateInstance($t, @($g))
$val = $t.GetProperty("Value").GetValue($inst)
if ($val -ne $g) {
    Write-Host "FAIL: Nullable guid value mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
