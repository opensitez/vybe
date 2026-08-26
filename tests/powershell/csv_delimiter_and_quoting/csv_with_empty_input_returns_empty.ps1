# vybe-test: powershell/csv_delimiter_and_quoting/csv_with_empty_input_returns_empty
$rows = @("" | ConvertFrom-Csv)
if ($rows.Length -ne 0) {
    Write-Host "FAIL: Empty CSV input should return empty"
    exit 1
}
Write-Host "PASS"
exit 0
