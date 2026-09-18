# vybe-test: powershell/format_table_engine/table_multiple_calculated_columns
$items = @([pscustomobject]@{ BaseRate = 100; TaxRate = 0.20 })

$col1 = @{ Label = "TaxAmount"; Expression = { $_.BaseRate * $_.TaxRate } }
$col2 = @{ Label = "GrandTotal"; Expression = { $_.BaseRate + ($_.BaseRate * $_.TaxRate) } }

$output = $items | Format-Table BaseRate, $col1, $col2 | Out-String

if (-not ($output -match "TaxAmount" -and $output -match "GrandTotal")) {
    Write-Host "FAIL: calculated column headers missing: $output"
    exit 1
}

if (-not ($output -match "20" -and $output -match "120")) {
    Write-Host "FAIL: calculated values 20 and 120 missing: $output"
    exit 1
}

Write-Host "PASS"
exit 0
