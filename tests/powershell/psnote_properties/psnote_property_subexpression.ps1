# vybe-test: powershell/psnote_properties/psnote_property_subexpression
$o = [pscustomobject]@{}
$o | Add-Member -NotePropertyName "Count" -NotePropertyValue 42
$msg = "Count: $( $o.Count )"
if ($msg -ne "Count: 42") {
    Write-Host "FAIL: NoteProperty subexpression expected 'Count: 42', got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
