# vybe-test: powershell/csv_type_coercion_on_import/explicit_decimal_parse_after_import
$csv = @"
Product,Price
Coffee,3.95
"@
$row = $csv | ConvertFrom-Csv
$price = [decimal]::Parse($row.Price, [System.Globalization.CultureInfo]::InvariantCulture)
if ($price -ne [decimal]3.95) {
    Write-Host "FAIL: Explicit Decimal parse after import failed, got $price"
    exit 1
}
Write-Host "PASS"
exit 0
