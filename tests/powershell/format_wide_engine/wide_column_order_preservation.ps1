# vybe-test: powershell/format_wide_engine/wide_column_order_preservation
$letters = @("FirstItem", "SecondItem", "ThirdItem") | ForEach-Object {
    [pscustomobject]@{ Item = $_ }
}

# In Format-Wide -Column 1, items must appear in exact input order
$output = $letters | Format-Wide -Property Item -Column 1 | Out-String

$firstIndex  = $output.IndexOf("FirstItem")
$secondIndex = $output.IndexOf("SecondItem")
$thirdIndex  = $output.IndexOf("ThirdItem")

if ($firstIndex -lt 0 -or $secondIndex -lt 0 -or $thirdIndex -lt 0) {
    Write-Host "FAIL: items missing from wide output: $output"
    exit 1
}

if (-not ($firstIndex -lt $secondIndex -and $secondIndex -lt $thirdIndex)) {
    Write-Host "FAIL: order not preserved in wide view, indices: First=$firstIndex, Second=$secondIndex, Third=$thirdIndex"
    exit 1
}

Write-Host "PASS"
exit 0
