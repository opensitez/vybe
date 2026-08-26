# vybe-test: powershell/csv_type_coercion_on_import/type_coercion_via_select_object_calculated_property
$csv = @"
Item,Cost
A,15
B,25
"@
$typed = @($csv | ConvertFrom-Csv | Select-Object Item, @{ N = "CostInt"; E = { [int]$_.Cost } })
if ($typed[0].CostInt -ne 15 -or $typed[0].CostInt -isnot [int]) {
    Write-Host "FAIL: Type coercion via calculated properties failed"
    exit 1
}
Write-Host "PASS"
exit 0
