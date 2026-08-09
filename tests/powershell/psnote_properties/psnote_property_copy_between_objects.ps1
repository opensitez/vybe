# vybe-test: powershell/psnote_properties/psnote_property_copy_between_objects
$o1 = [pscustomobject]@{ SourceProp = "Data" }
$o2 = [pscustomobject]@{}
foreach ($p in $o1.psobject.Properties) {
    if ($p.MemberType -eq "NoteProperty") {
        $o2 | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value
    }
}
if ($o2.SourceProp -ne "Data") {
    Write-Host "FAIL: NoteProperty copy between objects expected SourceProp=Data"
    exit 1
}
Write-Host "PASS"
exit 0
