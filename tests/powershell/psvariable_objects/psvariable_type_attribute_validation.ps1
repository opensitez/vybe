# vybe-test: powershell/psvariable_objects/psvariable_type_attribute_validation
$v = [System.Management.Automation.PSVariable]::new("TypedVar", "100")
$v.Attributes.Add([System.Management.Automation.PSTypeNameAttribute]::new("System.String"))
if ($v.Attributes.Count -ne 1) {
    Write-Host "FAIL: PSVariable PSTypeNameAttribute validation missing"
    exit 1
}
Write-Host "PASS"
exit 0
