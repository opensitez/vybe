# vybe-test: powershell/json_array_vs_single_object/string_array_json_parsing
$json = '["apple", "banana", "cherry"]'
$arr = @(ConvertFrom-Json -InputObject $json)
if ($arr.Length -ne 3 -or $arr[1] -ne "banana") {
    Write-Host "FAIL: String array JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
