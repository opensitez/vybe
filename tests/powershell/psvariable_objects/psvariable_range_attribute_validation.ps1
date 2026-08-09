# vybe-test: powershell/psvariable_objects/psvariable_range_attribute_validation
$v = [System.Management.Automation.PSVariable]::new("RangedVar", 5)
$v.Attributes.Add([System.Management.Automation.ValidateRangeAttribute]::new(1, 10))
if ($v.Attributes.Count -ne 1) {
    Write-Host "FAIL: ValidateRangeAttribute addition to PSVariable failed"
    exit 1
}
Write-Host "PASS"
exit 0
