# vybe-test: powershell/psalias_properties/psalias_property_remove
$obj = [pscustomobject]@{ Real = 1 }
$obj | Add-Member -MemberType AliasProperty -Name "TempAlias" -Value "Real"
$obj.psobject.Properties.Remove("TempAlias")
if ($obj.psobject.Properties["TempAlias"] -ne $null) {
    Write-Host "FAIL: AliasProperty removal failed"
    exit 1
}
Write-Host "PASS"
exit 0
