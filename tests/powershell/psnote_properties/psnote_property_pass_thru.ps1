# vybe-test: powershell/psnote_properties/psnote_property_pass_thru
$obj = [pscustomobject]@{ Id = 1 }
$returned = $obj | Add-Member -NotePropertyName "Status" -NotePropertyValue "OK" -PassThru
if ($returned.Status -ne "OK" -or $returned.Id -ne 1) {
    Write-Host "FAIL: Add-Member -PassThru expected returned object with Status=OK"
    exit 1
}
Write-Host "PASS"
exit 0
