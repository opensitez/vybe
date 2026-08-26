# vybe-test: powershell/csv_type_coercion_on_import/type_coercion_in_where_object_filter
$csv = @"
User,Age
Alice,30
Bob,17
Charlie,25
"@
$adults = @($csv | ConvertFrom-Csv | Where-Object { [int]$_.Age -ge 18 })
if ($adults.Length -ne 2 -or $adults[0].User -ne "Alice" -or $adults[1].User -ne "Charlie") {
    Write-Host "FAIL: Type coercion in Where-Object filter failed"
    exit 1
}
Write-Host "PASS"
exit 0
