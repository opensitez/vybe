# vybe-test: powershell/psnote_properties/psnote_property_type_check
$obj = [pscustomobject]@{}
$prop = [System.Management.Automation.PSNoteProperty]::new("Prop", "Val")
if (-not ($prop -is [System.Management.Automation.PSNoteProperty])) {
    Write-Host "FAIL: object is not [PSNoteProperty]"
    exit 1
}
Write-Host "PASS"
exit 0
