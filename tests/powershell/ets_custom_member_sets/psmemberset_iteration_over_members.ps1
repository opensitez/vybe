# vybe-test: powershell/ets_custom_member_sets/psmemberset_iteration_over_members
$obj = [pscustomobject]@{ Id = 100 }
$obj | Add-Member -NotePropertyName "Prop1" -NotePropertyValue "Val1"
if ($obj.Prop1 -ne "Val1" -or $obj.Id -ne 100) {
    Write-Host "FAIL: Member access failed"
    exit 1
}
Write-Host "PASS"
exit 0
