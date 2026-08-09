# vybe-test: powershell/psalias_properties/psalias_property_basic
$obj = [pscustomobject]@{ RealName = "TargetValue" }
$obj | Add-Member -MemberType AliasProperty -Name "NickName" -Value "RealName"
if ($obj.NickName -ne "TargetValue") {
    Write-Host "FAIL: AliasProperty NickName expected 'TargetValue', got '$($obj.NickName)'"
    exit 1
}
Write-Host "PASS"
exit 0
