# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_with_guid_roundtrip
$g = [guid]::NewGuid()
$orig = [pscustomobject]@{ Id = $g }
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Id -ne $g.ToString()) {
    Write-Host "FAIL: GUID property roundtrip in custom object failed"
    exit 1
}
Write-Host "PASS"
exit 0
