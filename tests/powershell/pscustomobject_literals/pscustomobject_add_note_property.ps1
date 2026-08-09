# vybe-test: powershell/pscustomobject_literals/pscustomobject_add_note_property
$obj = [pscustomobject]@{ Base = 1 }
$obj | Add-Member -NotePropertyName "Extra" -NotePropertyValue 2
if ($obj.Extra -ne 2) {
    Write-Host "FAIL: Add-Member expected Extra=2, got $($obj.Extra)"
    exit 1
}
Write-Host "PASS"
exit 0
