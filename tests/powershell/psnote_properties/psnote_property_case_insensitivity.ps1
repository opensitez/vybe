# vybe-test: powershell/psnote_properties/psnote_property_case_insensitivity
$obj = [pscustomobject]@{}
$obj | Add-Member -NotePropertyName "CamelCase" -NotePropertyValue "CaseData"
if ($obj.camelcase -ne "CaseData") {
    Write-Host "FAIL: case-insensitive NoteProperty access expected CaseData, got '$($obj.camelcase)'"
    exit 1
}
Write-Host "PASS"
exit 0
