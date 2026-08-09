# vybe-test: powershell/type_converters/type_converter_property_conversion
class Config {
    [version]$Version = "1.0.0.0"
}
$cfg = [Config]::new()
if ($cfg.Version.Major -ne 1) {
    Write-Host "FAIL: class property type converter expected Version Major=1"
    exit 1
}
Write-Host "PASS"
exit 0
