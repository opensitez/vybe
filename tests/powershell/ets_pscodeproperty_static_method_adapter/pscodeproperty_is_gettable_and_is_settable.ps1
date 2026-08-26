# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_is_gettable_and_is_settable
class SettableCheck {
    static [string]GetX([psobject]$i) { return "x" }
    static [void]SetX([psobject]$i, [string]$v) {}
}
$getter = [SettableCheck].GetMethod("GetX")
$setter = [SettableCheck].GetMethod("SetX")
$roProp = [System.Management.Automation.PSCodeProperty]::new("RO", $getter)
$rwProp = [System.Management.Automation.PSCodeProperty]::new("RW", $getter, $setter)
if (-not $roProp.IsGettable -or $roProp.IsSettable -or -not $rwProp.IsSettable) {
    Write-Host "FAIL: PSCodeProperty IsGettable / IsSettable failed"
    exit 1
}
Write-Host "PASS"
exit 0
