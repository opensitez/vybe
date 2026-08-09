# vybe-test: powershell/type_converters/type_converter_pipeline_coercion
$res = @("10", "20", "30") | ForEach-Object { [int]$_ }
if ($res[0] -ne 10 -or $res[2] -ne 30 -or -not ($res[0] -is [int])) {
    Write-Host "FAIL: pipeline element type converter expected int 10, 20, 30"
    exit 1
}
Write-Host "PASS"
exit 0
