# vybe-test: powershell/ets_psaliasproperty_over_nested_objects/alias_property_appears_in_get_member
$obj = [pscustomobject]@{ Num = 42 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSAliasProperty]::new("AltNum", "Num"))
$members = @($obj | Get-Member | Where-Object { $_.MemberType -eq "AliasProperty" })
if ($members.Length -ne 1 -or $members[0].Name -ne "AltNum") {
    Write-Host "FAIL: PSAliasProperty in Get-Member failed"
    exit 1
}
Write-Host "PASS"
exit 0
