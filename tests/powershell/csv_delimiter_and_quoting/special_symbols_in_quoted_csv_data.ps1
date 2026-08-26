# vybe-test: powershell/csv_delimiter_and_quoting/special_symbols_in_quoted_csv_data
$csv = @"
Symbol,Val
"#!@#$%^&*()_+",100
"@
$rows = @($csv | ConvertFrom-Csv)
if ($rows[0].Symbol -ne "#!@#$%^&*()_+") {
    Write-Host "FAIL: Special symbols in CSV failed, got '$($rows[0].Symbol)'"
    exit 1
}
Write-Host "PASS"
exit 0
