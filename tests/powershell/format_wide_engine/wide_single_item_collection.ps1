# vybe-test: powershell/format_wide_engine/wide_single_item_collection
$single = @([pscustomobject]@{ Tag = "SoloElement" })

# Single-item collections format properly without formatting artifacts
$output = $single | Format-Wide -Property Tag -Column 2 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "SoloElement")) {
    Write-Host "FAIL: single-element wide formatting missing element: $output"
    exit 1
}

Write-Host "PASS"
exit 0
