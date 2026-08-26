# vybe-test: powershell/json_array_vs_single_object/boolean_array_json_parsing
$json = '[true, false, true]'
$arr = @(ConvertFrom-Json -InputObject $json)
if ($arr.Length -ne 3 -or $arr[0] -ne $true -or $arr[1] -ne $false) {
    Write-Host "FAIL: Boolean array JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
