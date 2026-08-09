# vybe-test: powershell/psnote_properties/psnote_property_scriptblock_value
$obj = [pscustomobject]@{}
$obj | Add-Member -NotePropertyName "Code" -NotePropertyValue { param($a) $a * 5 }
$res = &($obj.Code) 4
if ($res -ne 20) {
    Write-Host "FAIL: NoteProperty scriptblock execution expected 20, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
