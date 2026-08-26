# vybe-test: powershell/classes_property_attributes/property_validation_with_type_conversion
class PortHolder {
    [ValidateRange(1, 65535)][int]$Port
}
$ph = [PortHolder]::new()
$ph.Port = "8080" # coerced string to int
if ($ph.Port -ne 8080) {
    Write-Host "FAIL: Property validation with coerced type failed"
    exit 1
}
Write-Host "PASS"
exit 0
