# vybe-test: powershell/csv_header_manipulation/headers_only_csv_returns_empty_rows
$csv = "Col1,Col2,Col3"
$rows = @($csv | ConvertFrom-Csv)
if ($rows.Length -ne 0) {
    Write-Host "FAIL: Header-only CSV should produce 0 data rows, got $($rows.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
