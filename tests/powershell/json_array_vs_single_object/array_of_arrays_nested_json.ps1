# vybe-test: powershell/json_array_vs_single_object/array_of_arrays_nested_json
$json = '[[1, 2], [3, 4]]'
$matrix = @(ConvertFrom-Json -InputObject $json)
if ($matrix[0][0] -ne 1 -or $matrix[1][1] -ne 4) {
    Write-Host "FAIL: Array of arrays JSON failed"
    exit 1
}
Write-Host "PASS"
exit 0
