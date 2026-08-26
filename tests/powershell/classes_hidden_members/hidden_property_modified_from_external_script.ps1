# vybe-test: powershell/classes_hidden_members/hidden_property_modified_from_external_script
class DirectAccess {
    hidden [string]$Val = "initial"
}
$d = [DirectAccess]::new()
$d.Val = "external_mutation"
if ($d.Val -ne "external_mutation") {
    Write-Host "FAIL: Direct assignment to hidden member failed"
    exit 1
}
Write-Host "PASS"
exit 0
