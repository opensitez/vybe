# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_membertype_check
$obj = [pscustomobject]@{ X = 10 }
$m = $obj.PSObject.Properties["X"]
if ($m.MemberType -ne [System.Management.Automation.PSMemberTypes]::NoteProperty) {
    Write-Host "FAIL: PSNoteProperty MemberType check failed"
    exit 1
}
Write-Host "PASS"
exit 0
