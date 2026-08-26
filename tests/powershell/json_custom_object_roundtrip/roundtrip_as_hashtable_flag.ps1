# vybe-test: powershell/json_custom_object_roundtrip/roundtrip_as_hashtable_flag
$orig = @{ A = 10; B = 20 }
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json -AsHashtable
if ($recovered -isnot [hashtable] -or $recovered["A"] -ne 10 -or $recovered["B"] -ne 20) {
    Write-Host "FAIL: ConvertFrom-Json -AsHashtable failed"
    exit 1
}
Write-Host "PASS"
exit 0
