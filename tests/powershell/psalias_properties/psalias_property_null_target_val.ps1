# vybe-test: powershell/psalias_properties/psalias_property_null_target_val
$obj = [pscustomobject]@{ NullableField = $null }
$obj | Add-Member -MemberType AliasProperty -Name "AliasNull" -Value "NullableField"
if ($obj.AliasNull -ne $null) {
    Write-Host "FAIL: AliasProperty for null field expected null, got $($obj.AliasNull)"
    exit 1
}
Write-Host "PASS"
exit 0
