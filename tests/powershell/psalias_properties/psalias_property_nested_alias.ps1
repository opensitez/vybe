# vybe-test: powershell/psalias_properties/psalias_property_nested_alias
$obj = [pscustomobject]@{ Origin = "DeepValue" }
$obj | Add-Member -MemberType AliasProperty -Name "AliasA" -Value "Origin"
$obj | Add-Member -MemberType AliasProperty -Name "AliasB" -Value "AliasA"
if ($obj.AliasB -ne "DeepValue") {
    Write-Host "FAIL: nested AliasProperty expected DeepValue, got '$($obj.AliasB)'"
    exit 1
}
Write-Host "PASS"
exit 0
