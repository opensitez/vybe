# vybe-test: powershell/csv_header_manipulation/custom_header_parameter_supplied
$csv = @"
Alice,30,Engineer
Bob,25,Designer
"@
$rows = @($csv | ConvertFrom-Csv -Header "Name", "Age", "Role")
if ($rows.Length -ne 2 -or $rows[0].Name -ne "Alice" -or $rows[1].Role -ne "Designer") {
    Write-Host "FAIL: Custom Header parameter import failed"
    exit 1
}
Write-Host "PASS"
exit 0
