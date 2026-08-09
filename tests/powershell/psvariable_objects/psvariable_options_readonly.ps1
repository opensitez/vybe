# vybe-test: powershell/psvariable_objects/psvariable_options_readonly
$v = [System.Management.Automation.PSVariable]::new("ROVar", "Immutable", [System.Management.Automation.ScopedItemOptions]::ReadOnly)
if ($v.Options -band [System.Management.Automation.ScopedItemOptions]::ReadOnly -eq 0) {
    Write-Host "FAIL: PSVariable ReadOnly option flag expected"
    exit 1
}
Write-Host "PASS"
exit 0
