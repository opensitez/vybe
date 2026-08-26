# vybe-test: powershell/classes_property_attributes/validaterange_on_double_property
class TempControl {
    [ValidateRange(-40.0, 100.0)][double]$Temp
}
$tc = [TempControl]::new()
$tc.Temp = 36.6
$caught = $false
try {
    $tc.Temp = 150.0
} catch {
    $caught = $true
}
if ($tc.Temp -ne 36.6 -or -not $caught) {
    Write-Host "FAIL: ValidateRange on double property failed"
    exit 1
}
Write-Host "PASS"
exit 0
