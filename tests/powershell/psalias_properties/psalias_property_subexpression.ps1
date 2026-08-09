# vybe-test: powershell/psalias_properties/psalias_property_subexpression
$obj = [pscustomobject]@{ DetailedName = "VybeCore" }
$obj | Add-Member -MemberType AliasProperty -Name "Name" -Value "DetailedName"
$msg = "Name: $( $obj.Name )"
if ($msg -ne "Name: VybeCore") {
    Write-Host "FAIL: AliasProperty in subexpression expected 'Name: VybeCore', got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
