# vybe-test: powershell/format_table_engine/table_hide_table_headers
$data = @(
    [pscustomobject]@{ ItemName = "Gadget"; Price = 100 },
    [pscustomobject]@{ ItemName = "Widget"; Price = 200 }
)

# -HideTableHeaders suppresses the column names and the dashed divider row
$output = $data | Format-Table -Property ItemName, Price -HideTableHeaders | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# The data rows must be present
if (-not ($output -match "Gadget" -and $output -match "Widget")) {
    Write-Host "FAIL: data rows missing from output: $output"
    exit 1
}

# The dashed divider line (e.g. '--------') must NOT be present
if ($output -match "--------" -or $output -match "========") {
    Write-Host "FAIL: divider row appeared despite -HideTableHeaders: $output"
    exit 1
}

Write-Host "PASS"
exit 0
