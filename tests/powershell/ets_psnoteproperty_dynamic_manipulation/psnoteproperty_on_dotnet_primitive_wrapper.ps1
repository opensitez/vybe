# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_on_dotnet_primitive_wrapper
$obj = [pscustomobject]@{ BaseNum = 42 }
$obj.PSObject.Properties.Add([System.Management.Automation.PSNoteProperty]::new("Extra", "Val"))
if ($obj.Extra -ne "Val") {
    Write-Host "FAIL: NoteProperty failed"
    exit 1
}
Write-Host "PASS"
exit 0
