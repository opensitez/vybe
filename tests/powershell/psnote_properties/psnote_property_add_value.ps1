# vybe-test: powershell/psnote_properties/psnote_property_add_value
$obj = [pscustomobject]@{ Base = 1 }
$obj.psobject.Properties.Add([System.Management.Automation.PSNoteProperty]::new("NoteVal", 100))
if ($obj.NoteVal -ne 100) {
    Write-Host "FAIL: PSNoteProperty object creation expected 100, got $($obj.NoteVal)"
    exit 1
}
Write-Host "PASS"
exit 0
