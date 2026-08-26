# vybe-test: powershell/ets_psscriptproperty_getter_and_setter/psscriptproperty_is_gettable_and_is_settable
$ro = [System.Management.Automation.PSScriptProperty]::new("RO", { 1 })
$rw = [System.Management.Automation.PSScriptProperty]::new("RW", { 1 }, { param($v) })
if (-not $ro.IsGettable -or $ro.IsSettable -or -not $rw.IsSettable) {
    Write-Host "FAIL: PSScriptProperty IsGettable / IsSettable failed"
    exit 1
}
Write-Host "PASS"
exit 0
