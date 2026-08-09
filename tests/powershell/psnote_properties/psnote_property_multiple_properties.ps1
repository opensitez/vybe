# vybe-test: powershell/psnote_properties/psnote_property_multiple_properties
$obj = [pscustomobject]@{}
$obj | Add-Member -NotePropertyMembers @{ P1 = "V1"; P2 = "V2" }
if ($obj.P1 -ne "V1" -or $obj.P2 -ne "V2") {
    Write-Host "FAIL: NotePropertyMembers hashtable expected P1=V1, P2=V2"
    exit 1
}
Write-Host "PASS"
exit 0
