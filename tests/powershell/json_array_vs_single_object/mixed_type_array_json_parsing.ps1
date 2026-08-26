# vybe-test: powershell/json_array_vs_single_object/mixed_type_array_json_parsing
$json = '[1, "text", true, null]'
$arr = @(ConvertFrom-Json -InputObject $json)
if ($arr.Length -ne 4 -or $arr[0] -ne 1 -or $arr[1] -ne "text" -or $arr[2] -ne $true -or $arr[3] -ne $null) {
    Write-Host "FAIL: Mixed type array JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
