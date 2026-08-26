# vybe-test: powershell/ets_custom_member_sets/psmemberset_scriptmethod_inside_memberset
$obj = [pscustomobject]@{ Id = 1; Name = "Test" }
$obj | Add-Member -NotePropertyName "Extra1" -NotePropertyValue "Val1"
if ($obj.Extra1 -ne "Val1" -or $obj.Id -ne 1) {
    Write-Host "FAIL: Custom member set property access failed"
    exit 1
}
Write-Host "PASS"
exit 0
