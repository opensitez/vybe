# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_dynamic_name_from_variable
$fieldName = "DynamicField"
$obj = [pscustomobject]@{}
$obj | Add-Member -NotePropertyName $fieldName -NotePropertyValue 999
if ($obj.$fieldName -ne 999 -or $obj.DynamicField -ne 999) {
    Write-Host "FAIL: Dynamic PSNoteProperty name from variable failed"
    exit 1
}
Write-Host "PASS"
exit 0
