# vybe-test: powershell/classes_property_attributes/validation_attribute_on_static_property
class StaticValidator {
    [ValidateRange(1, 100)] static [int]$Limit = 10
}
[StaticValidator]::Limit = 50
$caught = $false
try {
    [StaticValidator]::Limit = 200
} catch {
    $caught = $true
}
if ([StaticValidator]::Limit -ne 50 -or -not $caught) {
    Write-Host "FAIL: Validation attribute on static property failed"
    exit 1
}
Write-Host "PASS"
exit 0
