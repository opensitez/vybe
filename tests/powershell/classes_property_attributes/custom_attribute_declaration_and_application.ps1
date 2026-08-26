# vybe-test: powershell/classes_property_attributes/custom_attribute_declaration_and_application
class ValidatedPropDemo {
    [ValidateSet("Active", "Inactive")][string]$Status = "Active"
}
$inst = [ValidatedPropDemo]::new()
$inst.Status = "Inactive"
$caught = $false
try {
    $inst.Status = "BadStatus"
} catch {
    $caught = $true
}
if ($inst.Status -ne "Inactive" -or -not $caught) {
    Write-Host "FAIL: Class property validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
