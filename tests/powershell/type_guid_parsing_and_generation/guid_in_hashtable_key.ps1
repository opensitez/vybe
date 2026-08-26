# vybe-test: powershell/type_guid_parsing_and_generation/guid_in_hashtable_key
$g = [guid]::NewGuid()
$ht = @{ $g = "val" }
if ($ht[$g] -ne "val") {
    Write-Host "FAIL: GUID as hashtable key lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
