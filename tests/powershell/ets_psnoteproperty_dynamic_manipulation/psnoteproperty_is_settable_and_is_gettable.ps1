# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_is_settable_and_is_gettable
$note = [System.Management.Automation.PSNoteProperty]::new("Prop", "Val")
if (-not $note.IsGettable -or -not $note.IsSettable) {
    Write-Host "FAIL: PSNoteProperty IsGettable / IsSettable failed"
    exit 1
}
Write-Host "PASS"
exit 0
