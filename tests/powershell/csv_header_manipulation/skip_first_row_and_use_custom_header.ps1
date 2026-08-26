# vybe-test: powershell/csv_header_manipulation/skip_first_row_and_use_custom_header
$csv = @"
OldHeader1,OldHeader2
Val1,Val2
Val3,Val4
"@
# Skip original header line, use custom headers
$dataLines = $csv -split "`r?`n" | Select-Object -Skip 1
$rows = @($dataLines | ConvertFrom-Csv -Header "Col1", "Col2")
if ($rows.Length -ne 2 -or $rows[0].Col1 -ne "Val1" -or $rows[1].Col2 -ne "Val4") {
    Write-Host "FAIL: Skip first row and custom header failed, got count $($rows.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
