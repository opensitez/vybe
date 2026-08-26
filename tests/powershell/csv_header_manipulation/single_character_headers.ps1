# vybe-test: powershell/csv_header_manipulation/single_character_headers
$csv = @"
X,Y,Z
10,20,30
"@
$row = $csv | ConvertFrom-Csv
if ($row.X -ne "10" -or $row.Y -ne "20" -or $row.Z -ne "30") {
    Write-Host "FAIL: Single character headers failed"
    exit 1
}
Write-Host "PASS"
exit 0
