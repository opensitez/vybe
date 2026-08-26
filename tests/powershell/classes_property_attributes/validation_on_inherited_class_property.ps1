# vybe-test: powershell/classes_property_attributes/validation_on_inherited_class_property
class ParentPropClass {
    [ValidateRange(1, 100)][int]$Capacity = 50
}
class ChildPropClass : ParentPropClass {}
$cp = [ChildPropClass]::new()
$cp.Capacity = 75
$caught = $false
try {
    $cp.Capacity = 150
} catch {
    $caught = $true
}
if ($cp.Capacity -ne 75 -or -not $caught) {
    Write-Host "FAIL: Validation on inherited property failed"
    exit 1
}
Write-Host "PASS"
exit 0
