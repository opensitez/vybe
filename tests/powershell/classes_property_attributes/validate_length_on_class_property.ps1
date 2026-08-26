# vybe-test: powershell/classes_property_attributes/validate_length_on_class_property
class FixedCodeClass {
    [ValidateLength(2, 5)][string]$Code = "ABC"
}
$fcc = [FixedCodeClass]::new()
$fcc.Code = "XYZW"
$caught = $false
try {
    $fcc.Code = "Toolongcodehere"
} catch {
    $caught = $true
}
if ($fcc.Code -ne "XYZW" -or -not $caught) {
    Write-Host "FAIL: ValidateLength on class property failed"
    exit 1
}
Write-Host "PASS"
exit 0
