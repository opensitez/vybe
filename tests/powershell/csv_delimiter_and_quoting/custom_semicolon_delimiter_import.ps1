# vybe-test: powershell/csv_delimiter_and_quoting/custom_semicolon_delimiter_import
$csv = @"
Name;Age;Role
Alice;30;Engineer
Bob;25;Designer
"@
$rows = @($csv | ConvertFrom-Csv -Delimiter ';')
if ($rows.Length -ne 2 -or $rows[0].Name -ne "Alice" -or $rows[0].Age -ne "30" -or $rows[1].Role -ne "Designer") {
    Write-Host "FAIL: Semicolon delimiter import failed"
    exit 1
}
Write-Host "PASS"
exit 0
