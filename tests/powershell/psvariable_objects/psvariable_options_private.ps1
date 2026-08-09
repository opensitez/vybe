# vybe-test: powershell/psvariable_objects/psvariable_options_private
$v = [System.Management.Automation.PSVariable]::new("PrivVar", "PrivateVal", [System.Management.Automation.ScopedItemOptions]::Private)
if ($v.Options -band [System.Management.Automation.ScopedItemOptions]::Private -eq 0) {
    Write-Host "FAIL: PSVariable Private option flag expected"
    exit 1
}
Write-Host "PASS"
exit 0
