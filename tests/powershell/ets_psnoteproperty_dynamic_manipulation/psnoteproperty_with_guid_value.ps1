# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_with_guid_value
$g = [guid]::NewGuid()
$obj = [pscustomobject]@{ Id = $g }
if ($obj.Id -ne $g -or $obj.Id -isnot [guid]) {
    Write-Host "FAIL: PSNoteProperty with GUID value failed"
    exit 1
}
Write-Host "PASS"
exit 0
