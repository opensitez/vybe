# vybe-test: powershell/json_array_vs_single_object/single_item_array_coerced_via_at_syntax
$json = '{"item":1}'
$arr = @($json | ConvertFrom-Json)
if ($arr.Length -ne 1 -or $arr[0].item -ne 1) {
    Write-Host "FAIL: Single item coerced via @() array syntax failed"
    exit 1
}
Write-Host "PASS"
exit 0
