# vybe-test: powershell/psalias_properties/psalias_property_write_through
$obj = [pscustomobject]@{ Underlying = "Initial" }
$obj | Add-Member -MemberType AliasProperty -Name "Facade" -Value "Underlying"
$obj.Facade = "ModifiedViaFacade"
if ($obj.Underlying -ne "ModifiedViaFacade") {
    Write-Host "FAIL: write through AliasProperty expected Underlying='ModifiedViaFacade', got '$($obj.Underlying)'"
    exit 1
}
Write-Host "PASS"
exit 0
