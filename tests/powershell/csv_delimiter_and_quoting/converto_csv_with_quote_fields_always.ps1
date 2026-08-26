# vybe-test: powershell/csv_delimiter_and_quoting/converto_csv_with_quote_fields_always
$items = @([pscustomobject]@{ Num = 123; Text = "Hello" })
$csv = @($items | ConvertTo-Csv -NoTypeInformation)
if (-not $csv[1].Contains('"123"') -and -not $csv[1].Contains('"Hello"')) {
    Write-Host "FAIL: Quoting of fields in ConvertTo-Csv failed, got '$($csv[1])'"
    exit 1
}
Write-Host "PASS"
exit 0
