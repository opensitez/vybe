# vybe-test: powershell/json_array_vs_single_object/single_element_array_json_returns_array_or_enumerable
$json = '[{"id":42}]'
$res = @(ConvertFrom-Json -InputObject $json)
if ($res.Length -ne 1 -or $res[0].id -ne 42) {
    Write-Host "FAIL: Single element array JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
