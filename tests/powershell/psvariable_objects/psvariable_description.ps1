# vybe-test: powershell/psvariable_objects/psvariable_description
$v = [System.Management.Automation.PSVariable]::new("Described", 1, "None", "Custom variable description")
if ($v.Description -ne "Custom variable description") {
    Write-Host "FAIL: PSVariable Description expected 'Custom variable description', got '$($v.Description)'"
    exit 1
}
Write-Host "PASS"
exit 0
