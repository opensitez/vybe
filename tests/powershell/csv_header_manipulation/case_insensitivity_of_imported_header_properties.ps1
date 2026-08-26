# vybe-test: powershell/csv_header_manipulation/case_insensitivity_of_imported_header_properties
$csv = @"
UserName,EmailAddress
alice,alice@example.com
"@
$rows = @($csv | ConvertFrom-Csv)
if ($rows[0].username -ne "alice" -or $rows[0].USERNAME -ne "alice") {
    Write-Host "FAIL: Case-insensitivity of CSV header properties failed"
    exit 1
}
Write-Host "PASS"
exit 0
