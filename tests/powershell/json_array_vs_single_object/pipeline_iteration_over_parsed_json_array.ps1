# vybe-test: powershell/json_array_vs_single_object/pipeline_iteration_over_parsed_json_array
$json = '[{"v":10},{"v":20},{"v":30}]'
$sum = 0
ConvertFrom-Json -InputObject $json | ForEach-Object { $sum += $_.v }
if ($sum -ne 60) {
    Write-Host "FAIL: Pipeline iteration over JSON array failed, sum=$sum"
    exit 1
}
Write-Host "PASS"
exit 0
