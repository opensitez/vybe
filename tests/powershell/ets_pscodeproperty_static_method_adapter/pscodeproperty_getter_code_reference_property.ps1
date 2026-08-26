# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_getter_code_reference_property
class RefCheck {
    static [string]GetRef([psobject]$inst) { return "ref" }
}
$m = [RefCheck].GetMethod("GetRef")
$prop = [System.Management.Automation.PSCodeProperty]::new("RefProp", $m)
if ($prop.GetterCodeReference -ne $m -or $prop.SetterCodeReference -ne $null) {
    Write-Host "FAIL: GetterCodeReference / SetterCodeReference check failed"
    exit 1
}
Write-Host "PASS"
exit 0
