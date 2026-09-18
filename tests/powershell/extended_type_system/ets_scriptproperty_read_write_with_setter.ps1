# vybe-test: powershell/extended_type_system/ets_scriptproperty_read_write_with_setter
# Update-TypeData supports read-write ScriptProperty with -SecondValue acting as the setter
class DeviceState { [int]$rawTemp = 20 }

Update-TypeData -TypeName DeviceState -MemberType ScriptProperty -MemberName TempF -Value {
    $this.rawTemp * 9/5 + 32
} -SecondValue {
    param($val)
    $this.rawTemp = [int](($val - 32) * 5/9)
} -Force

$dev = [DeviceState]::new()

if ($dev.TempF -ne 68) {
    Write-Host "FAIL: getter failed, expected 68, got: $($dev.TempF)"
    exit 1
}

$dev.TempF = 212

if ($dev.rawTemp -ne 100) {
    Write-Host "FAIL: setter failed, expected rawTemp 100, got: $($dev.rawTemp)"
    exit 1
}

Write-Host "PASS"
exit 0
