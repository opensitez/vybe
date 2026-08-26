# vybe-test: powershell/enums_flags_attribute/enum_flags_in_class_property
[System.FlagsAttribute()]
enum DeviceFeatures {
    Wifi = 1
    Bluetooth = 2
    Gps = 4
}
class Device {
    [DeviceFeatures]$Features
    Device([DeviceFeatures]$f) { $this.Features = $f }
}
$d = [Device]::new([DeviceFeatures]::Wifi -bor [DeviceFeatures]::Gps)
if (-not $d.Features.HasFlag([DeviceFeatures]::Wifi) -or $d.Features.HasFlag([DeviceFeatures]::Bluetooth)) {
    Write-Host "FAIL: Flags enum in class property failed"
    exit 1
}
Write-Host "PASS"
exit 0
