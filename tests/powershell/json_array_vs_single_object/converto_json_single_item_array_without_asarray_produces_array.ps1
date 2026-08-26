# vybe-test: powershell/json_array_vs_single_object/converto_json_single_item_array_without_asarray_produces_array
$singleArr = @("only")
$json = $singleArr | ConvertTo-Json
# An array with 1 element serialized without -AsArray may serialize as array or object depending on pipeline unrolling
$recovered = @($json | ConvertFrom-Json)
if ($recovered[0] -ne "only") {
    Write-Host "FAIL: Single item array roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
