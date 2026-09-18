# vybe-test: powershell/format_list_engine/list_calculated_property_expression
$order = [pscustomobject]@{ UnitPrice = 25; Quantity = 4 }

# Format-List supports calculated property hashtables with Label and Expression
$calcProp = @{ Label = "TotalAmount"; Expression = { $_.UnitPrice * $_.Quantity } }
$output = $order | Format-List $calcProp | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# The calculated label and evaluated value (25*4 = 100) must be rendered
if (-not ($output -match "TotalAmount\s*:\s*100")) {
    Write-Host "FAIL: calculated property failed to render in list view: $output"
    exit 1
}

Write-Host "PASS"
exit 0
