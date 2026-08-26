# vybe-test: powershell/json_array_vs_single_object/single_string_json_scalar
$json = '"hello world"'
$val = ConvertFrom-Json -InputObject $json
if ($val -ne "hello world") {
    Write-Host "FAIL: Single string scalar JSON failed, got '$val'"
    exit 1
}
Write-Host "PASS"
exit 0
