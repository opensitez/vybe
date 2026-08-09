# vybe-test: powershell/psalias_properties/psalias_property_add_member_cmdlet
$obj = [pscustomobject]@{ Original = "Data" }
Add-Member -InputObject $obj -MemberType AliasProperty -Name "Alias" -Value "Original"
if ($obj.Alias -ne "Data") {
    Write-Host "FAIL: Add-Member AliasProperty expected Alias='Data', got '$($obj.Alias)'"
    exit 1
}
Write-Host "PASS"
exit 0
