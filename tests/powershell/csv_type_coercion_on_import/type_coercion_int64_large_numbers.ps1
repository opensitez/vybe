# vybe-test: powershell/csv_type_coercion_on_import/type_coercion_int64_large_numbers
$csv = @"
Bytes
9000000000000
"@
$row = $csv | ConvertFrom-Csv
$bytes = [int64]::Parse($row.Bytes)
if ($bytes -ne 9000000000000) {
    Write-Host "FAIL: Int64 large number coercion failed"
    exit 1
}
Write-Host "PASS"
exit 0
