# vybe-test: powershell/psvariable_objects/psvariable_attributes_list
$v = [System.Management.Automation.PSVariable]::new("AttributedVar", "Data")
$v.Attributes.Add([System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new())
if ($v.Attributes.Count -ne 1) {
    Write-Host "FAIL: PSVariable Attributes count expected 1"
    exit 1
}
Write-Host "PASS"
exit 0
