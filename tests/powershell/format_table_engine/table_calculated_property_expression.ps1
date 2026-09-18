# vybe-test: powershell/format_table_engine/table_calculated_property_expression
$orders = @(
    [pscustomobject]@{ Product = "Keyboard"; UnitPrice = 40; Quantity = 2 },
    [pscustomobject]@{ Product = "Mouse";    UnitPrice = 15; Quantity = 3 }
)

# Calculated property hashtable with Label and Expression
$calcCol = @{ Label = "SubTotal"; Expression = { $_.UnitPrice * $_.Quantity } }
$output = $orders | Format-Table Product, $calcCol | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: Format-Table output was null"
    exit 1
}

# The calculated column header must be present
if (-not ($output -match "SubTotal")) {
    Write-Host "FAIL: calculated column header 'SubTotal' missing from output: $output"
    exit 1
}

# Calculated values 80 (40*2) and 45 (15*3) must appear
if (-not ($output -match "80" -and $output -match "45")) {
    Write-Host "FAIL: calculated values (80 and 45) missing from output: $output"
    exit 1
}

Write-Host "PASS"
exit 0
