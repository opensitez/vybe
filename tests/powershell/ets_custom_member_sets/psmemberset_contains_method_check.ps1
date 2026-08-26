# vybe-test: powershell/ets_custom_member_sets/psmemberset_contains_method_check
$set = [System.Management.Automation.PSMemberSet]::new("CheckSet")
$set.Members.Add([System.Management.Automation.PSNoteProperty]::new("Item1", "Val"))
if ($set.Members["Item1"] -eq $null -or $set.Members["NonExistent"] -ne $null) {
    Write-Host "FAIL: PSMemberSet Members indexer check failed"
    exit 1
}
Write-Host "PASS"
exit 0
