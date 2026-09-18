# vybe-test: powershell/extended_type_system/ets_extend_custom_powershell_class
# PowerShell user-defined classes can be extended via Update-TypeData
class InventoryBox { [string]$Code = "BOX10" }

Update-TypeData -TypeName InventoryBox -MemberType ScriptProperty -MemberName LowerCode -Value {
    $this.Code.ToLower()
} -Force

$box = [InventoryBox]::new()

if ($box.LowerCode -ne "box10") {
    Write-Host "FAIL: custom class ETS property failed, got: '$($box.LowerCode)'"
    exit 1
}

Write-Host "PASS"
exit 0
