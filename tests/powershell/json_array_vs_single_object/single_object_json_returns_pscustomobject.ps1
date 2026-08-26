# vybe-test: powershell/json_array_vs_single_object/single_object_json_returns_pscustomobject
$json = '{"id":1,"name":"single"}'
$obj = ConvertFrom-Json -InputObject $json
if ($obj -isnot [pscustomobject] -or $obj.id -ne 1) {
    Write-Host "FAIL: Single object JSON should produce PSCustomObject"
    exit 1
}
Write-Host "PASS"
exit 0
