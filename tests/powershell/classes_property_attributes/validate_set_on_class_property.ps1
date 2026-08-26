# vybe-test: powershell/classes_property_attributes/validate_set_on_class_property
class AccountClass {
    [ValidateSet("Active", "Inactive")][string]$Status = "Active"
}
$ac = [AccountClass]::new()
$ac.Status = "Inactive"
$caught = $false
try {
    $ac.Status = "InvalidStatus"
} catch {
    $caught = $true
}
if ($ac.Status -ne "Inactive" -or -not $caught) {
    Write-Host "FAIL: ValidateSet on class property failed"
    exit 1
}
Write-Host "PASS"
exit 0
