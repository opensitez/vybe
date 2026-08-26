# vybe-test: powershell/json_array_vs_single_object/array_with_null_elements
$json = '[1, null, 3]'
$arr = @(ConvertFrom-Json -InputObject $json)
if ($arr.Length -ne 3 -or $arr[1] -ne $null) {
    Write-Host "FAIL: Array with null elements failed"
    exit 1
}
Write-Host "PASS"
exit 0
