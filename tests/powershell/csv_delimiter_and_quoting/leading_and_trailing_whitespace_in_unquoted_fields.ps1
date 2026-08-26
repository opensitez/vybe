# vybe-test: powershell/csv_delimiter_and_quoting/leading_and_trailing_whitespace_in_unquoted_fields
$csv = "Col1,Col2`n  spaced  ,  trimmed  "
$rows = @($csv | ConvertFrom-Csv)
if ($rows[0].Col1.Trim() -ne "spaced") {
    Write-Host "FAIL: Unquoted whitespace preservation failed"
    exit 1
}
Write-Host "PASS"
exit 0
