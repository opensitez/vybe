# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_case_insensitivity
class CaseCode { static [int]GetVal([psobject]$i) { return 42 } }
$obj = [pscustomobject]@{ X = 1 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("TheAnswer", [CaseCode].GetMethod("GetVal")))
if ($obj.theanswer -ne 42 -or $obj.THEANSWER -ne 42) {
    Write-Host "FAIL: PSCodeProperty case-insensitivity failed"
    exit 1
}
Write-Host "PASS"
exit 0
