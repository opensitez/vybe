# vybe-test: powershell/csv_type_coercion_on_import/type_coercion_in_measure_object
$csv = @"
Score
10
20
30
"@
$m = $csv | ConvertFrom-Csv | ForEach-Object { [int]$_.Score } | Measure-Object -Sum -Average
if ($m.Sum -ne 60 -or $m.Average -ne 20.0) {
    Write-Host "FAIL: Measure-Object on coerced CSV scores failed"
    exit 1
}
Write-Host "PASS"
exit 0
