# vybe-test: powershell/psvariable_objects/psvariable_value_mutation
$v = [System.Management.Automation.PSVariable]::new("MutVar", 10)
$v.Value = 20
if ($v.Value -ne 20) {
    Write-Host "FAIL: PSVariable Value mutation expected 20, got $($v.Value)"
    exit 1
}
Write-Host "PASS"
exit 0
