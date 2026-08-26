# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_getter_invoking_static_method
class CodeHelper {
    static [string]GetFormattedId([psobject]$instance) {
        return "ID-" + $instance.Id
    }
}
$obj = [pscustomobject]@{ Id = 1234 }
$getterMethod = [CodeHelper].GetMethod("GetFormattedId")
$prop = [System.Management.Automation.PSCodeProperty]::new("FormattedId", $getterMethod)
$obj.PSObject.Properties.Add($prop)
if ($obj.FormattedId -ne "ID-1234") {
    Write-Host "FAIL: PSCodeProperty getter failed, got '$($obj.FormattedId)'"
    exit 1
}
Write-Host "PASS"
exit 0
