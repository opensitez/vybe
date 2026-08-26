# vybe-test: powershell/csv_type_coercion_on_import/explicit_double_parse_after_import
$csv = @"
Item,Rate
USD,1.25
"@
$row = $csv | ConvertFrom-Csv
$rate = [double]::Parse($row.Rate, [System.Globalization.CultureInfo]::InvariantCulture)
if ($rate -ne 1.25) {
    Write-Host "FAIL: Explicit double parse after import failed, got $rate"
    exit 1
}
Write-Host "PASS"
exit 0
