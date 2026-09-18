# vybe-test: powershell/extended_type_system/ets_remove_typedata_unregisters_members
# Remove-TypeData deregisters dynamic ETS type definitions
class TempRemovableType { [string]$Name = "Initial" }

Update-TypeData -TypeName TempRemovableType -MemberType ScriptProperty -MemberName Upper -Value { $this.Name.ToUpper() } -Force
$inst = [TempRemovableType]::new()

if ($inst.Upper -ne "INITIAL") {
    Write-Host "FAIL: initial ETS member call failed"
    exit 1
}

$remData = Get-TypeData -TypeName TempRemovableType
Remove-TypeData -TypeData $remData

# After Remove-TypeData, accessing the property returns $null or throws under strict mode
$valAfterRemove = $inst.Upper

if ($null -ne $valAfterRemove) {
    Write-Host "FAIL: member still returned value after Remove-TypeData: '$valAfterRemove'"
    exit 1
}

Write-Host "PASS"
exit 0
