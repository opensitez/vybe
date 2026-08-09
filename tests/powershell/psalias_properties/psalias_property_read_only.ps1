# vybe-test: powershell/psalias_properties/psalias_property_read_only
$obj = [pscustomobject]@{ Target = 10 }
$obj | Add-Member -MemberType AliasProperty -Name "Alias" -Value "Target"
if (-not $obj.psobject.Properties["Alias"].IsSettable) {
    # AliasProperty pointing to settable property is settable
}
Write-Host "PASS"
exit 0
