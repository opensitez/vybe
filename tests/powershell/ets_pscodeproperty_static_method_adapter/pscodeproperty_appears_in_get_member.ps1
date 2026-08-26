# vybe-test: powershell/ets_pscodeproperty_static_method_adapter/pscodeproperty_appears_in_get_member
class GmCode {
    static [string]GetGm([psobject]$i) { return "gm" }
}
$obj = [pscustomobject]@{ A = 1 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSCodeProperty]::new("GmProp", [GmCode].GetMethod("GetGm")))
$members = @($obj | Get-Member | Where-Object { $_.MemberType -eq "CodeProperty" })
if ($members.Length -ne 1 -or $members[0].Name -ne "GmProp") {
    Write-Host "FAIL: PSCodeProperty in Get-Member failed"
    exit 1
}
Write-Host "PASS"
exit 0
