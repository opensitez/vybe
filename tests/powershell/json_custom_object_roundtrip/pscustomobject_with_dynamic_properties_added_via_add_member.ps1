# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_with_dynamic_properties_added_via_add_member
$orig = [pscustomobject]@{ Base = 1 }
$orig | Add-Member -NotePropertyName "Dynamic" -NotePropertyValue "Added"
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Base -ne 1 -or $recovered.Dynamic -ne "Added") {
    Write-Host "FAIL: Dynamic properties added via Add-Member roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
