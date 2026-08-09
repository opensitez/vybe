# vybe-test: powershell/psalias_properties/psalias_property_multiple_aliases
$obj = [pscustomobject]@{ BaseData = 55 }
$obj | Add-Member -MemberType AliasProperty -Name "Alias1" -Value "BaseData"
$obj | Add-Member -MemberType AliasProperty -Name "Alias2" -Value "BaseData"
if ($obj.Alias1 -ne 55 -or $obj.Alias2 -ne 55) {
    Write-Host "FAIL: multiple AliasProperties pointing to BaseData expected 55"
    exit 1
}
Write-Host "PASS"
exit 0
