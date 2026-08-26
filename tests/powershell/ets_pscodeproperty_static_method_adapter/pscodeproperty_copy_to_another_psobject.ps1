# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_copy_to_another_psobject
class CopyCode { static [string]GetHi([psobject]$i) { return "Hi" } }
$m = [CopyCode].GetMethod("GetHi")
$p1 = [pscustomobject]@{ Name = "A" }
$p2 = [pscustomobject]@{ Name = "B" }
$p1.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("Greeting", $m))
$p2.PSObject.Properties.Add($p1.PSObject.Properties["Greeting"].Copy())
if ($p2.Greeting -ne "Hi") {
    Write-Host "FAIL: PSCodeProperty Copy() failed"
    exit 1
}
Write-Host "PASS"
exit 0
