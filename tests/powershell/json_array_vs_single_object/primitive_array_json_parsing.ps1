# vybe-test: powershell/json_array_vs_single_object/primitive_array_json_parsing
$json = '[10, 20, 30, 40]'
$arr = @(ConvertFrom-Json -InputObject $json)
if ($arr.Length -ne 4 -or $arr[0] -ne 10 -or $arr[3] -ne 40) {
    Write-Host "FAIL: Primitive array JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
