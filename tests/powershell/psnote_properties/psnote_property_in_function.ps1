# vybe-test: powershell/psnote_properties/psnote_property_in_function
function Decorate-Object($target, [string]$name, $val) {
    $target | Add-Member -NotePropertyName $name -NotePropertyValue $val
}
$o = [pscustomobject]@{}
Decorate-Object $o "Decorated" "Yes"
if ($o.Decorated -ne "Yes") {
    Write-Host "FAIL: NoteProperty added in function expected Decorated=Yes"
    exit 1
}
Write-Host "PASS"
exit 0
