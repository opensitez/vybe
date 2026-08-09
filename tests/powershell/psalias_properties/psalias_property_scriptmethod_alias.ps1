# vybe-test: powershell/psalias_properties/psalias_property_scriptmethod_alias
$obj = [pscustomobject]@{ Base = 10 }
$obj | Add-Member -MemberType ScriptMethod -Name "Compute" -Value { $this.Base * 2 }
$obj | Add-Member -MemberType AliasProperty -Name "CalcAlias" -Value "Compute"
# AliasProperty pointing to method produces callable alias or evaluates member
if ($obj.psobject.Properties["CalcAlias"] -eq $null) {
    Write-Host "FAIL: AliasProperty registered for method missing"
    exit 1
}
Write-Host "PASS"
exit 0
