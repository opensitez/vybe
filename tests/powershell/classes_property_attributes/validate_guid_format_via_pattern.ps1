# vybe-test: powershell/classes_property_attributes/validate_guid_format_via_pattern
class GuidFormatCheck {
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')][string]$GuidStr
}
$gfc = [GuidFormatCheck]::new()
$gfc.GuidStr = "12345678-1234-1234-1234-123456789abc"
$caught = $false
try {
    $gfc.GuidStr = "not-a-guid"
} catch {
    $caught = $true
}
if ($gfc.GuidStr -ne "12345678-1234-1234-1234-123456789abc" -or -not $caught) {
    Write-Host "FAIL: GUID pattern validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
