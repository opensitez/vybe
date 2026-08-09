# vybe-test: powershell/psvariable_objects/psvariable_options_constant
$v = [System.Management.Automation.PSVariable]::new("ConstVar", "ConstantVal", [System.Management.Automation.ScopedItemOptions]::Constant)
if ($v.Options -band [System.Management.Automation.ScopedItemOptions]::Constant -eq 0) {
    Write-Host "FAIL: PSVariable Constant option flag expected"
    exit 1
}
Write-Host "PASS"
exit 0
