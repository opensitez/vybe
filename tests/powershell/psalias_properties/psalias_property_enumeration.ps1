# vybe-test: powershell/psalias_properties/psalias_property_enumeration
$obj = [pscustomobject]@{ Real = 1 }
$obj | Add-Member -MemberType AliasProperty -Name "Alt" -Value "Real"
$aliases = $obj.psobject.Properties | Where-Object { $_.MemberType -eq "AliasProperty" }
if ($aliases.Count -ne 1 -or $aliases[0].Name -ne "Alt") {
    Write-Host "FAIL: AliasProperty enumeration expected Alt"
    exit 1
}
Write-Host "PASS"
exit 0
