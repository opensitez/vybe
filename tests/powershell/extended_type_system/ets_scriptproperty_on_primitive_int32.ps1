# vybe-test: powershell/extended_type_system/ets_scriptproperty_on_primitive_int32
# Primitive value types (like System.Int32) can be extended dynamically via ETS
Update-TypeData -TypeName System.Int32 -MemberType ScriptProperty -MemberName Squared -Value { $this * $this } -Force

$val = (7).Squared

if ($val -ne 49) {
    Write-Host "FAIL: expected (7).Squared to be 49, got: $val"
    exit 1
}

Write-Host "PASS"
exit 0
