# vybe-test: powershell/format_table_engine/table_calculated_property_alignment_right
$data = @([pscustomobject]@{ Metric = "Hits"; Value = 12345 })

# Alignment key in calculated property hashtable formats alignment as Right
$rightCol = @{ Label = "Value"; Expression = { $_.Value }; Alignment = "Right" }
$output = $data | Format-Table Metric, $rightCol | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "Value" -and $output -match "12345")) {
    Write-Host "FAIL: right-aligned column content missing: $output"
    exit 1
}

Write-Host "PASS"
exit 0
