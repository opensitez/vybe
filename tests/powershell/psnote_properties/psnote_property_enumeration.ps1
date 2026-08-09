# vybe-test: powershell/psnote_properties/psnote_property_enumeration
$obj = [pscustomobject]@{ A = 1; B = 2 }
$notes = $obj.psobject.Properties | Where-Object { $_.MemberType -eq "NoteProperty" }
if ($notes.Count -ne 2) {
    Write-Host "FAIL: NoteProperty count expected 2, got $($notes.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
