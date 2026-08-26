# vybe-test: powershell/classes_property_attributes/validate_range_on_class_property
class BoundedClass {
    [ValidateRange(1, 100)][int]$Score = 50
}
$bc = [BoundedClass]::new()
$bc.Score = 99
$caught = $false
try {
    $bc.Score = 150
} catch {
    $caught = $true
}
if ($bc.Score -ne 99 -or -not $caught) {
    Write-Host "FAIL: ValidateRange on class property failed"
    exit 1
}
Write-Host "PASS"
exit 0
