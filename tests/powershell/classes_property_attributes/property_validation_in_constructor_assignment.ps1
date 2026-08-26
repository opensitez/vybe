# vybe-test: powershell/classes_property_attributes/property_validation_in_constructor_assignment
class StrictConstruct {
    [ValidateRange(1, 10)][int]$Rank
    StrictConstruct([int]$r) { $this.Rank = $r }
}
$ok = [StrictConstruct]::new(5)
$caught = $false
try {
    $bad = [StrictConstruct]::new(99)
} catch {
    $caught = $true
}
if ($ok.Rank -ne 5 -or -not $caught) {
    Write-Host "FAIL: Constructor property validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
