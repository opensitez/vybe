# vybe-test: powershell/format_list_engine/list_property_order_preservation
$obj = [pscustomobject]@{ Zebra = "Last"; Apple = "First" }

# In Format-List, properties must appear in the exact order requested by the -Property argument
$output = $obj | Format-List -Property Apple, Zebra | Out-String

$appleIndex = $output.IndexOf("Apple")
$zebraIndex = $output.IndexOf("Zebra")

if ($appleIndex -lt 0 -or $zebraIndex -lt 0) {
    Write-Host "FAIL: properties missing from output: $output"
    exit 1
}

if (-not ($appleIndex -lt $zebraIndex)) {
    Write-Host "FAIL: property order not preserved, expected Apple before Zebra, got Apple=$appleIndex, Zebra=$zebraIndex"
    exit 1
}

Write-Host "PASS"
exit 0
