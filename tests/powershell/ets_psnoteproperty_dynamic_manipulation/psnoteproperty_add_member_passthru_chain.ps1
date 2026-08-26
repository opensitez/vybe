# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_add_member_passthru_chain
$obj = [pscustomobject]@{} |
    Add-Member -NotePropertyName "A" -NotePropertyValue 1 -PassThru |
    Add-Member -NotePropertyName "B" -NotePropertyValue 2 -PassThru
if ($obj.A -ne 1 -or $obj.B -ne 2) {
    Write-Host "FAIL: Add-Member PassThru chaining failed"
    exit 1
}
Write-Host "PASS"
exit 0
