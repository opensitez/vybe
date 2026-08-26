# vybe-test: powershell/json_custom_object_roundtrip/ordered_dictionary_roundtrip
$orig = [ordered]@{
    First = 1
    Second = 2
    Third = 3
}
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.First -ne 1 -or $recovered.Second -ne 2 -or $recovered.Third -ne 3) {
    Write-Host "FAIL: Ordered dictionary roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
