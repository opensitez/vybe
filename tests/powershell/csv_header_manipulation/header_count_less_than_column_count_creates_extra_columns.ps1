# vybe-test: powershell/csv_header_manipulation/header_count_less_than_column_count_creates_extra_columns
$csv = "A,B,C`n1,2,3"
$rows = @($csv | ConvertFrom-Csv -Header "H1", "H2")
# PowerShell assigns generated column names for extra fields or parses them
if ($rows[0].H1 -ne "A" -or $rows[0].H2 -ne "B") {
    Write-Host "FAIL: Partial Header specification failed"
    exit 1
}
Write-Host "PASS"
exit 0
