# vybe-test: powershell/csv_type_coercion_on_import/type_coercion_in_sort_object_numeric_order
$csv = @"
File,Size
a,100
b,20
c,5
"@
# If sorted as strings: "100", "20", "5"
# If coerced to int: 5, 20, 100
$sorted = @($csv | ConvertFrom-Csv | Sort-Object { [int]$_.Size })
if ($sorted[0].File -ne "c" -or $sorted[1].File -ne "b" -or $sorted[2].File -ne "a") {
    Write-Host "FAIL: Numeric sort with type coercion failed"
    exit 1
}
Write-Host "PASS"
exit 0
