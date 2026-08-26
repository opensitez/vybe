# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_construction_and_read
$obj = [pscustomobject]@{}
$note = [System.Management.Automation.PSNoteProperty]::new("MyNote", "MyValue")
$obj.PSObject.Properties.Add($note)
if ($obj.MyNote -ne "MyValue") {
    Write-Host "FAIL: PSNoteProperty read failed, got '$($obj.MyNote)'"
    exit 1
}
Write-Host "PASS"
exit 0
