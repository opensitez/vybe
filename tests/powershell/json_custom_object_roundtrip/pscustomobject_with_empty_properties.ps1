# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_with_empty_properties
$orig = [pscustomobject]@{
    EmptyStr = ""
    EmptyArr = @()
    NullVal = $null
}
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.EmptyStr -ne "" -or $recovered.EmptyArr.Count -ne 0 -or $recovered.NullVal -ne $null) {
    Write-Host "FAIL: Empty properties roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
