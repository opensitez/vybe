# vybe-test: powershell/json_array_vs_single_object/empty_array_json_returns_empty_collection
$json = '[]'
$res = @(ConvertFrom-Json -InputObject $json)
if ($res.Length -ne 0) {
    Write-Host "FAIL: Empty array JSON should return 0 items, got $($res.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
