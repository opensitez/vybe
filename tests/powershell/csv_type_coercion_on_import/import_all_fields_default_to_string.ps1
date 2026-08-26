# vybe-test: powershell/csv_type_coercion_on_import/import_all_fields_default_to_string
$csv = @"
Id,Active,Price
100,True,49.99
"@
$row = $csv | ConvertFrom-Csv
if ($row.Id -isnot [string] -or $row.Active -isnot [string] -or $row.Price -isnot [string]) {
    Write-Host "FAIL: CSV fields should import as string by default"
    exit 1
}
Write-Host "PASS"
exit 0
