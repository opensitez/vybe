# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_removal_via_properties_remove
$obj = [pscustomobject]@{ A = 1; B = 2 }
$obj.PSObject.Properties.Remove("B")
if ($obj.B -ne $null -or $obj.A -ne 1 -or $obj.PSObject.Properties.Count -ne 1) {
    Write-Host "FAIL: PSNoteProperty removal failed"
    exit 1
}
Write-Host "PASS"
exit 0
