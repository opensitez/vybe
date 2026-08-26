# vybe-test: powershell/csv_header_manipulation/numeric_header_names
$csv = @"
"100","200"
A,B
"@
$row = $csv | ConvertFrom-Csv
if ($row."100" -ne "A" -or $row."200" -ne "B") {
    Write-Host "FAIL: Numeric header names failed"
    exit 1
}
Write-Host "PASS"
exit 0
